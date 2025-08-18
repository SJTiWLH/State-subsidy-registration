# 🚀 国补登记数据匹配工具

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
[![Open Source Love](https://badges.frapsoft.com/os/v1/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badges/)

## 📖 项目介绍

本项目是一款高效的数据处理工具，专为国家补贴登记设计。通过自动化流程实现：

- 🔄 **智能数据比对**：自动匹配抖音国补订单与3C网站商品信息
- 🧩 **精准字段匹配**：智能映射3C商品名称到订单数据
- 📊 **规格自动识别**：根据预设规则提取商品规格信息
- 💾 **一键输出结果**：生成符合国补登记要求的规范化数据文件

## ✨ 功能亮点

- ⚡ **高效自动化**：分钟级处理万条数据，效率提升90%
- 🧠 **智能规则引擎**：支持自定义匹配规则，灵活适应不同需求
- 📈 **可视化报告**：生成清晰的数据处理报告，问题一目了然
- 📁 **格式兼容性强**：完美支持Excel文件处理
- 🛡️ **数据安全保障**：处理过程不修改原始数据

## 🛠️ 安装要求

### 环境依赖
- **Python 3.7+** (推荐使用最新稳定版)
- **依赖库**：

## 🚀 安装步骤


# 1. 克隆仓库到本地
```bash
https://github.com/SJTiWLH/State-subsidy-registration.git

```

# 2. 创建虚拟环境（推荐）
```bash
python -m venv .venv
```

# 3. 激活虚拟环境
## Windows系统
```bash
.venv\Scripts\activate
```
## Mac/Linux系统
```bash
source .venv/bin/activate
```

# 4. 安装依赖包
```bash
pip install -r requirements.txt
```

## 🚦 使用方法

### 1️⃣ 准备输入文件

| 文件类型 | 文件名 | 说明 |
|----------|--------|------|
| 抖音国补订单 | `3c商品名表格` | 包含订单号、商品名称等关键字段，存档到该文件夹中 |
| 3C商品数据 | `抖音表格` | 包含3C编号、商品名称、规格等信息 |
| 企业库存数量 | `企业库存数量.xlsx`|从在线表格上下载最新的并保存在根目录|

### 2️⃣ 运行程序

```bash
python main.py
```

```bash
./run.sh
```

### 3️⃣ 参数配置 (config.ini)

```bash
[MATCHING_RULES]
# 商品名称匹配规则 (支持：精确匹配、模糊匹配、包含关键字)
product_name_match = 包含关键字

# 规格匹配规则 (支持正则表达式)
spec_match = (\d+[GgTt]?B?|\d+寸|\d+英寸)
```

### 4️⃣ 输出结果

| 文件 | 说明 |
|------|------|
| `matched_results.xlsx` | 匹配完成的完整数据 |
| `unmatched_items.csv` | 未匹配商品记录 |
| `report.html` | 可视化处理报告 |

## 📂 文件结构

```bash
national-subsidy-matcher/
├── README.md
├── main.py                # 主程序入口
├── config.ini             # 配置文件
├── requirements.txt       # 依赖库列表
├── data/                  # 数据目录
│   ├── input/             # 输入文件存放处
│   └── output/            # 处理结果输出位置
├── utils/                 # 核心功能模块
│   ├── matcher.py         # 智能匹配引擎
│   ├── reporter.py        # 报告生成器
│   └── validator.py       # 数据验证工具
└── tests/                 # 单元测试
    ├── test_matcher.py
    └── test_data/
```

## ⚠️ 注意事项

1. **数据格式要求**
   - ✅ 确保抖音订单包含：订单号、商品名称、购买数量
   - ✅ 确保3C数据包含：3C编号、标准商品名称、规格参数

2. **数据安全**
   - 🔒 首次使用前务必备份原始数据
   - 📂 建议在`data/input/`目录下操作

3. **匹配问题排查**
   - 🔍 商品名称不一致？尝试修改匹配规则为"模糊匹配"
   - 📏 规格描述不规范？添加新的正则表达式规则
   - ❌ 3C数据库不完整？更新3C数据源文件

## 🤝 参与贡献

欢迎任何形式的贡献！请遵循以下流程：

1. Fork 项目仓库
2. 创建您的功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

## 📜 许可证

本项目采用 [MIT License](LICENSE) - 详情请参阅 LICENSE 文件。

---

<p align="center">
  让数据处理更简单 · 让补贴登记更高效
</p>
