import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# 设置页面配置
st.set_page_config(
    page_title="商品调货建议系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🛍️ 商品调货建议系统")
st.markdown("---")

# 侧边栏配置
st.sidebar.header("📁 数据上传")
uploaded_file = st.sidebar.file_uploader("上传Excel文件", type=["xlsx"])

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 系统设置")
run_analysis = st.sidebar.button("🚀 运行分析", type="primary")

# 数据预处理函数
def preprocess_data(df):
    notes = []
    
    # 强制转换Article为12位文本
    df['Article'] = df['Article'].astype(str).str.zfill(12)
    
    # 处理整数字段
    integer_columns = ['SaSa Net Stock', 'Pending Received', 'Safety Stock', 
                      'Last Month Sold Qty', 'MTD Sold Qty']
    
    for col in integer_columns:
        if col in df.columns:
            # 记录转换前的非数字值
            original_values = df[col].copy()
            # 转换为数值并处理NaN
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # 查找转换失败的值
            invalid_rows = df[df[col].isna() & original_values.notna()]
            for idx in invalid_rows.index:
                notes.append(f"行 {idx+2}: {col} 列值 '{original_values[idx]}' 转换为0")
            
            # 填充NaN为0
            df[col] = df[col].fillna(0).astype(int)
            
            # 处理异常值
            below_zero = df[col] < 0
            above_limit = df[col] > 100000
            
            if below_zero.any():
                for idx in df[below_zero].index:
                    notes.append(f"行 {idx+2}: {col} 值小于0，已修正为0")
                df.loc[below_zero, col] = 0
                
            if above_limit.any():
                for idx in df[above_limit].index:
                    notes.append(f"行 {idx+2}: {col} 值超出范围，已修正为100000")
                df.loc[above_limit, col] = 100000
    
    # 处理文本字段
    text_columns = ['OM', 'RP Type', 'Site']
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
    
    return df, notes

# 核心业务逻辑函数
def calculate_transfer_suggestions(df):
    # 定义有效销量
    df['Effective Sold Qty'] = np.where(
        df['Last Month Sold Qty'] > 0, 
        df['Last Month Sold Qty'], 
        df['MTD Sold Qty']
    )
    
    # 创建结果列表
    transfer_suggestions = []
    summary_stats = {
        'total_transfers': 0,
        'total_qty': 0,
        'articles_count': set(),
        'oms_count': set()
    }
    
    # 按Article和OM分组
    grouped = df.groupby(['Article', 'OM'])
    
    for (article, om), group in grouped:
        summary_stats['articles_count'].add(article)
        summary_stats['oms_count'].add(om)
        
        # 识别转出店铺
        sources = []
        
        # 优先级1: ND类型转出
        nd_sources = group[group['RP Type'] == 'ND'].copy()
        nd_sources['Transferable Qty'] = nd_sources['SaSa Net Stock']
        for _, row in nd_sources.iterrows():
            sources.append({
                'Site': row['Site'],
                'Type': 'ND',
                'Qty': row['Transferable Qty'],
                'Priority': 1,
                'Row': row  # 保存行数据以获取Product Desc等信息
            })
        
        # 优先级2: RF类型过剰转出
        rf_group = group[group['RP Type'] == 'RF'].copy()
        if not rf_group.empty:
            rf_group['Available Qty'] = rf_group['SaSa Net Stock'] + rf_group['Pending Received']
            rf_group['Excess'] = rf_group['Available Qty'] - rf_group['Safety Stock']
            # 找出销量不是最高的店铺
            max_sold = rf_group['Effective Sold Qty'].max() if not rf_group.empty else 0
            rf_sources = rf_group[(rf_group['Excess'] > 0) & (rf_group['Effective Sold Qty'] < max_sold)]
            
            for _, row in rf_sources.iterrows():
                sources.append({
                    'Site': row['Site'],
                    'Type': 'RF',
                    'Qty': row['Excess'],
                    'Priority': 2,
                    'Row': row  # 保存行数据
                })
        
        # 识别接收店铺
        destinations = []
        
        # 优先级1: 紧急缺货补货
        urgent_destinations = rf_group[
            (rf_group['SaSa Net Stock'] == 0) & 
            (rf_group['Effective Sold Qty'] > 0)
        ].copy()
        urgent_destinations['Needed Qty'] = urgent_destinations['Safety Stock']
        
        for _, row in urgent_destinations.iterrows():
            destinations.append({
                'Site': row['Site'],
                'Priority': 1,
                'Qty': row['Needed Qty'],
                'Type': 'Urgent',
                'Row': row  # 保存行数据
            })
        
        # 优先级2: 潜在缺货补货
        # 找出销量最高的店铺
        if not rf_group.empty:
            max_sold = rf_group['Effective Sold Qty'].max()
            potential_destinations = rf_group[
                ((rf_group['SaSa Net Stock'] + rf_group['Pending Received']) < rf_group['Safety Stock']) &
                (rf_group['Effective Sold Qty'] == max_sold)
            ].copy()
            potential_destinations['Needed Qty'] = (
                potential_destinations['Safety Stock'] - 
                (potential_destinations['SaSa Net Stock'] + potential_destinations['Pending Received'])
            )
            
            for _, row in potential_destinations.iterrows():
                destinations.append({
                    'Site': row['Site'],
                    'Priority': 2,
                    'Qty': row['Needed Qty'],
                    'Type': 'Potential',
                    'Row': row  # 保存行数据
                })
        
        # 执行匹配
        sources_sorted = sorted(sources, key=lambda x: (x['Priority'], -x['Qty']))
        destinations_sorted = sorted(destinations, key=lambda x: (x['Priority'], -x['Qty']))
        
        # 匹配逻辑
        for source in sources_sorted:
            if source['Qty'] <= 0:
                continue
                
            for dest in destinations_sorted:
                if dest['Qty'] <= 0:
                    continue
                    
                if source['Site'] == dest['Site']:
                    continue  # 不能自己调给自己
                    
                transfer_qty = min(source['Qty'], dest['Qty'])
                
                if transfer_qty > 0:
                    # 添加建议记录
                    product_desc = source['Row']['Product Desc'] if 'Product Desc' in source['Row'] else ''
                    
                    transfer_suggestions.append({
                        'Article': article,
                        'Product Desc': product_desc,
                        'OM': om,
                        'Transfer Site': source['Site'],
                        'Receive Site': dest['Site'],
                        'Transfer Qty': transfer_qty,
                        'Notes': f"转出类型: {source['Type']}, 接收优先级: {dest['Priority']}"
                    })
                    
                    # 更新统计
                    summary_stats['total_transfers'] += 1
                    summary_stats['total_qty'] += transfer_qty
                    
                    # 更新剩余数量
                    source['Qty'] -= transfer_qty
                    dest['Qty'] -= transfer_qty
    
    return transfer_suggestions, summary_stats

# 生成Excel报告
def generate_excel_report(transfer_suggestions, summary_stats):
    # 创建Excel文件
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 调货建议工作表
        if transfer_suggestions:
            df_recommendations = pd.DataFrame(transfer_suggestions)
            df_recommendations.to_excel(writer, sheet_name='调货建议', index=False)
        else:
            pd.DataFrame(columns=['Article', 'Product Desc', 'OM', 'Transfer Site', 
                                'Receive Site', 'Transfer Qty', 'Notes']).to_excel(
                                    writer, sheet_name='调货建议', index=False)
        
        # 统计摘要工作表
        summary_data = {
            '指标': ['总调货建议数量', '总调货件数', '涉及商品数量', '涉及OM数量'],
            '数值': [summary_stats['total_transfers'], 
                    summary_stats['total_qty'],
                    len(summary_stats['articles_count']),
                    len(summary_stats['oms_count'])]
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='统计摘要', index=False)
    
    output.seek(0)
    return output

# 主应用逻辑
if uploaded_file is not None:
    try:
        # 读取Excel文件
        df = pd.read_excel(uploaded_file, dtype={'Article': str})
        st.success("✅ Excel文件读取成功")
        
        # 显示原始数据信息
        with st.expander("看查看原始数据信息"):
            st.write(f"数据形状: {df.shape}")
            st.write("列名:", df.columns.tolist())
            st.dataframe(df.head())
        
        # 数据预处理
        processed_df, notes = preprocess_data(df)
        st.success("✅ 数据预处理完成")
        
        # 显示数据处理备注
        if notes:
            with st.expander("⚠️ 数据处理备注"):
                for note in notes:
                    st.warning(note)
        
        # 运行分析
        if run_analysis:
            with st.spinner("正在分析数据并生成调货建议..."):
                transfer_suggestions, summary_stats = calculate_transfer_suggestions(processed_df)
            
            # 显示结果
            st.header("📊 分析结果")
            
            # KPI指标
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("总调货建议数", summary_stats['total_transfers'])
            col2.metric("总调货件数", summary_stats['total_qty'])
            col3.metric("涉及商品数", len(summary_stats['articles_count']))
            col4.metric("涉及OM数", len(summary_stats['oms_count']))
            
            # 调货建议表格
            st.subheader("📋 调货建议详情")
            if transfer_suggestions:
                df_suggestions = pd.DataFrame(transfer_suggestions)
                st.dataframe(df_suggestions)
                
                # 下载按钮
                excel_data = generate_excel_report(transfer_suggestions, summary_stats)
                today = datetime.now().strftime("%Y%m%d")
                st.download_button(
                    label="📥 下载调货建议报告",
                    data=excel_data,
                    file_name=f"调货建议_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("未生成调货建议")
                
    except Exception as e:
        st.error(f"处理文件时出错: {str(e)}")
else:
    st.info("请在侧边栏上传Excel文件开始分析")
    
# 页脚
st.markdown("---")
st.caption("© 2025 调货建议系统 - 基于库存、销量和安全库存数据自动生成调货建议")