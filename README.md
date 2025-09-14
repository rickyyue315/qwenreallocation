# 商品调货建议系统

这是一个基于Streamlit构建的Web应用，可以根据库存、销量和安全库存数据自动生成跨店铺的商品调货建议。

## 功能特点

- 读取Excel格式的库存数据
- 根据预定义的业务规则自动计算调货建议
- 生成包含调货建议和统计摘要的Excel报告
- 提供友好的Web界面

## 安装和运行

1. 安装依赖：
   ```
   pip install -r requirements.txt
   ```

2. 运行应用：
   ```
   streamlit run app.py
   ```

## 使用说明

1. 在侧边栏上传包含库存数据的Excel文件
2. 点击"运行分析"按钮
3. 查看生成的调货建议
4. 下载Excel格式的调货建议报告

## 业务规则

### 数据预处理与验证

- Article字段强制转换为12位文本格式
- 整数字段(SaSa Net Stock, Pending Received等)处理异常值
- 文本字段(OM, RP Type, Site)处理缺失值
- 销量数据异常值校正（小于0设为0，大于100,000设为100,000）

### 调货规则

所有调货决策在相同Article和OM分组内进行：

1. 定义有效销量：优先使用Last Month Sold Qty，若为0或缺失则使用MTD Sold Qty
2. 转出规则：
   - 优先级1：RP Type为"ND"的店铺，可转数量为全部SaSa Net Stock
   - 优先级2：RP Type为"RF"且库存充足（SaSa Net Stock + Pending Received > Safety Stock）且销量不是最高值的店铺，可转数量为SaSa Net Stock + Pending Received - Safety Stock
3. 接收规则：
   - 优先级1：RP Type为"RF"且完全无库存（SaSa Net Stock = 0）且曾有销售记录（Effective Sold Qty > 0）的店铺，需求量为Safety Stock
   - 优先级2：RP Type为"RF"且库存不足（SaSa Net Stock + Pending Received < Safety Stock）且销量是最高值的店铺，需求量为Safety Stock - (SaSa Net Stock + Pending Received)

### 匹配与生成建议

1. 按Article和OM分组数据
2. 识别转出和接收候选店铺
3. 按优先级进行匹配（转出优先级 -> 接收优先级）
4. 每次匹配数量为min(转出方可转数量, 接收方需求量)
5. 生成调货建议记录

## 输出格式

生成的Excel报告包含两个工作表：

1. 调货建议：包含Article、Product Desc、OM、转出店铺、接收店铺、调货数量和备注
2. 统计摘要：包含总调货建议数量、总调货件数等关键指标