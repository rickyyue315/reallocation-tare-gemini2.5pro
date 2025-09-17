# 智能调货优化系统

基于Python开发的智能调货优化系统，专为零售库存管理设计。系统能够自动分析Excel库存数据，生成最优的跨店铺调货建议。

## 🚀 主要功能

### 1. 数据预处理与验证
- ✅ Article字段12位文本格式化
- ✅ 数值字段自动校正和异常值处理
- ✅ 缺失值智能填充
- ✅ 销量数据范围检查（0-100,000）

### 2. 核心业务逻辑
- 🔄 **ND类型转出**: 优先处理ND类型的库存转出
- 🔄 **RF类型转出**: 智能识别RF类型的过剩库存
- 🚨 **紧急缺货接收**: 零库存但有销售记录的店铺优先补货
- 📈 **潜在缺货接收**: 高销量店铺的库存优化

### 3. 智能匹配算法
- 🤖 Article+OM组合分组处理
- ⚡ 优先级驱动的匹配逻辑
- 📊 最优调货数量计算
- 🔄 实时状态更新

### 4. 质量保证体系
- ✅ 5项自动化质量检查
- 📋 完整的处理日志
- 🎯 数据准确性验证

### 5. 多格式输出
- 📊 Excel格式输出（调货建议 + 统计摘要）
- 📝 CSV格式导出
- 📈 可视化统计报表

## 🛠️ 技术栈

- **Python 3.8+**
- **Pandas** - 数据处理和分析
- **Streamlit** - Web用户界面
- **Openpyxl** - Excel文件处理
- **Logging** - 系统日志记录

## 📦 安装与运行

### 方式一：使用批处理文件（推荐）
1. 双击运行 `run_app.bat`
2. 系统自动安装依赖并启动Web界面

### 方式二：手动安装
```bash
# 安装依赖
pip install -r requirements.txt

# 启动Web界面
streamlit run web_interface.py

# 或者直接处理Excel文件
python transfer_system.py
```

## 📁 文件结构

```
智能调货优化系统/
├── transfer_system.py    # 核心调货逻辑
├── web_interface.py     # Web用户界面
├── run_app.bat          # Windows启动脚本
├── requirements.txt     # Python依赖包
├── README_CN.md        # 中文说明文档
└── sample_data.py      # 示例数据生成
```

## 📊 输入文件要求

### 必需字段
| 字段名 | 描述 | 格式 |
|--------|------|------|
| Article | 商品编号 | 文本 |
| RP Type | 类型标识 | ND/RF |
| Site | 店铺位置 | 文本 |
| OM | 组织单元 | 文本 |
| SaSa Net Stock | 当前库存 | 整数 |
| Safety Stock | 安全库存 | 整数 |

### 可选字段
| 字段名 | 描述 |
|--------|------|
| Last Month Sold Qty | 上月销量 |
| MTD Sold Qty | 本月至今销量 |
| Pending Received | 在途库存 |

## 🎯 输出成果

### 调货建议表
- Article (12位文本)
- OM (组织单元)
- Transfer Site (转出店铺)
- Receive Site (接收店铺)
- Transfer Qty (调货数量)
- Transfer Type (调货类型)
- Receive Priority (接收优先级)

### 统计摘要
- 📈 KPI关键指标
- 📊 按Article统计
- 📊 按OM统计
- 📊 调货类型分析
- 📊 接收优先级分析

## 🔧 自定义配置

系统支持以下参数调整：
- 安全库存系数调整
- 销量异常值阈值设置
- 输出文件格式选择
- 处理日志级别配置

## 📞 技术支持

如遇问题，请检查：
1. Python版本是否为3.8+
2. 所有依赖包是否安装成功
3. 输入文件格式是否符合要求
4. 系统日志中的错误信息

## 📄 许可证

本项目基于MIT许可证开源，可自由使用和修改。