# XlsxSearcher

Excel 子表搜索工具 - 快速查找 xlsx 文件中的子表名称

## 功能特性

- 🔍 **子表搜索**: 根据子表名称模糊搜索
- 📁 **文件搜索**: 根据文件名搜索
- 📊 **组合搜索**: 同时按子表名称和文件名搜索
- 📂 **一键打开**: 双击或按钮直接用 Excel 打开文件
- 🎯 **定位文件**: 在资源管理器/Finder 中定位并选中文件
- 📋 **复制路径**: 一键复制文件完整路径到剪贴板
- ⚡ **索引加速**: 首次扫描后建立 SQLite 索引，搜索毫秒级响应
- 🔄 **增量更新**: 重新扫描只更新有变化的文件

## 环境要求

- Python 3.8+
- macOS / Windows / Linux

### 依赖安装

```bash
pip install -r requirements.txt
```


## 使用方法

```bash
python main.py
```

### 操作流程

1. 点击 **"选择目录"** 按钮选择要扫描的文件夹
2. 程序会自动扫描目录下所有 xlsx 文件并建立索引
3. 在搜索框输入关键词（子表名称或文件名）进行搜索
4. 选中搜索结果，点击底部按钮执行操作：
   - **打开文件**: 用默认程序打开文件
   - **定位文件**: 在文件管理器中选中文件
   - **复制路径**: 复制文件路径到剪贴板

其他功能：
- **重新扫描**: 重新扫描当前选择的目录（只更新有变化的文件）
- **清空索引**: 清除所有已建立的索引数据

## 项目结构

```
XlsxSearcher/
├── main.py              # 程序入口
├── requirements.txt     # 依赖
├── core/
│   ├── indexer.py       # SQLite 索引管理
│   ├── scanner.py       # xlsx 文件扫描
│   └── searcher.py      # 搜索逻辑
├── gui/
│   └── app.py           # PyQt5 主界面
└── utils/
    └── file_utils.py    # 文件操作工具
```

## 打包发布

### macOS

```bash
# 已在依赖中包含
pyinstaller --onefile --windowed --name XlsxSearcher --icon=icon.icns main.py
```

生成的 `.app` 文件在 `dist` 目录下。

### Windows

```bash
# 已在依赖中包含
pyinstaller --onefile --windowed --name XlsxSearcher --icon=icon.ico main.py
```

生成的 `.exe` 文件在 `dist` 目录下。

### Linux

```bash
# 已在依赖中包含
pyinstaller --onefile --windowed --name XlsxSearcher main.py
```

生成的可执行文件在 `dist` 目录下。

## 数据存储

索引数据库保存在用户目录下：

- **所有平台**: `~/XlsxSearcher/index.db`

## 许可证

MIT License