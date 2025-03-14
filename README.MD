# TestViz - 电池测试数据可视化工具

## 项目概述
TestViz 是一个基于 Excel VBA 开发的电池测试数据可视化工具，专门用于处理和分析电池测试过程中的各项性能数据，并通过图表形式直观地展示测试结果。

## 主要功能

### 1. 循环寿命数据处理
- 处理并可视化容量保持率数据
- 处理并可视化能量保持率数据
- 支持多个电池数据的同时展示和对比

### 2. 中检数据处理
- 处理并可视化容量恢复率数据
- 处理并可视化 DCIR 增长率数据
- 支持多电池数据的对比分析

### 3. 数据可视化
- 自动生成专业的数据图表
- 支持自定义电池标识和颜色
- 图表包含完整的标题、坐标轴标签等信息

## 使用方法

1. **数据准备**
   - 确保数据文件包含 "Cycle Life" 和 "RPT of Cycle Life" 工作表
   - 数据格式需符合指定要求

2. **电池信息配置**
   - 在 "电池名字颜色" 表格中配置电池的名称和对应的显示颜色

3. **运行程序**
   - 运行主程序（Main 过程）即可自动处理数据并生成图表

## 系统要求

- Microsoft Excel（支持 VBA）
- 建议使用 Excel 2010 或更高版本

## 注意事项

- 运行过程中会暂时关闭屏幕更新和自动计算，以提高性能
- 请确保数据格式符合要求，否则可能导致处理错误
- 建议在处理大量数据时保存当前工作

## 错误处理

程序包含完整的错误处理机制：
- 数据格式错误会有相应提示
- 处理过程中的错误会被捕获并显示错误信息
- 关键步骤都有错误日志记录