# 多表格统一检索系统

> Multi-Table Search — 在本机把一个料号查遍任意多个 Excel 工作簿的每一个工作表

![platform](https://img.shields.io/badge/platform-Windows%2010%2B%20x64-3469f5)
![python](https://img.shields.io/badge/python-3.12-3469f5)
![license](https://img.shields.io/badge/license-MIT-4c9a6a)

![检索页](docs/screenshots/search.png)

## 这个工具解决什么问题

料号、型号、供应商报价散落在几个巨大的 Excel 里：一个文件几十个工作表，每个工作表
一套自己的表头，`Ctrl+F` 只能在工作表内找一个值，找到之后还得自己数行、自己对列。

这个工具把**任意多个** `.xlsx` 工作簿的所有工作表一次性建成索引。输入料号 → 返回所有
命中记录，每条都带着文件名、工作表、Excel 原始行号和**完整的一行数据**，并且能一键
用 WPS 打开并定位到单元格。

没有数量上限：3 个、30 个、300 个都行，随时可以再加一个，也可以单独移除。

## 主要特性

- **数量不设上限**：数据源页列出所有已索引工作簿，可随时添加、替换、移除
- **单元格级精确匹配**：逐单元格判断包含关系，不拼接相邻单元格，因此不会出现跨单元格的假命中
- **返回完整原始行**：重复表头按列字母区分（`供应商（H列）`），无表头列标注为 `未命名列`，右侧备注单独标为 `补充信息`
- **一键定位**：优先用 WPS COM 打开并跳到命中的单元格；失败时降级为打开文件并给你可复制的精确定位信息，绝不假装成功
- **导入是安全的**：先复制到暂存区、校验 SHA-256、解析、校验、最后单个事务提交；任何一步失败都保留旧索引和旧副本，并自动备份数据库
- **原件只读**：不修改、不移动、不上传任何 Excel
- **完全离线**：没有网络调用，不接入任何云端 API，不收集数据

## 界面

| 检索 | 数据源 |
| --- | --- |
| ![检索页](docs/screenshots/search.png) | ![数据源页](docs/screenshots/library.png) |

## 快速开始

### 安装包

Windows 10 1809（build 17763）或更新的 x64 系统，目标机需自行安装 WPS Office。
安装包在 [Releases](../../releases) 提供；也可以按 [BUILD_GUIDE.md](BUILD_GUIDE.md) 自行构建。

安装包自带 Python 运行时，目标机不需要装 Python、Node.js 或任何开发工具。

### 从源码运行

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m app.main
```

`--self-test` 会用合成数据在临时目录跑一遍真实解析、SQLite 和界面，不接触你的正式数据：

```powershell
.venv\Scripts\python.exe -m app.main --self-test
```

## 使用

1. 打开「数据源」→「添加工作簿」，选中一个或多个 `.xlsx`
2. 导入完成后回到「检索」，输入料号或关键词，按 Enter
3. 结果支持分页、复制单行、复制本页，以及「打开原始数据」

匹配规则是**逐单元格包含匹配**：大小写不敏感，只去掉首尾空格，不删除连字符、
不做模糊纠错、不解释数据。同一行多个单元格命中时只输出一次，但会列出全部命中位置。

## 数据放在哪

```
%LOCALAPPDATA%\MultiTableSearch\
├── database\search.db     索引数据库
├── workbooks\<版本>\      导入时的只读托管副本
├── backups\               每次导入前的事务一致备份
├── logs\app.log
└── settings\
```

程序安装目录和用户数据目录是分开的。卸载默认保留用户数据；要彻底删除请先退出程序，
再手动删除上面的目录。

备份恢复（先关闭正在运行的实例）：

```powershell
MultiTableSearch.exe --restore-backup "C:\...\backups\xxxx.db"
```

## 项目结构

```
app/
├── contracts.py          模块间唯一的数据与服务接口（INTERFACE_VERSION 2.0）
├── main.py               组合根、命令行参数、安装自检
├── core/
│   ├── database.py       SQLite 结构与备份恢复
│   ├── import_service.py 导入流水线（复制→解析→校验→事务提交）
│   ├── search_service.py 逐单元格包含检索与分页
│   └── result_formatter.py 纯文本结果格式化
├── excel_parser/         OOXML 解析：表头识别、合并单元格、公式缓存、空 fill 修复
├── windows/              用户数据路径、WPS COM 定位
└── ui/                   主窗口、检索页、数据源页、样式与后台任务
tests/                    72 个单元与界面测试（pytest + pytest-qt）
docs/                     架构、产品规格、UI 设计系统、解析说明、测试报告
```

界面规范、色板与偏离说明见 [docs/UI_DESIGN.md](docs/UI_DESIGN.md)；模块边界见
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。

## 开发

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest -q
```

测试默认使用合成工作簿并跑在 `offscreen` 平台，不依赖真实业务文件，也不需要显示器。

## 已知限制

- 只支持 `.xlsx`；`.xls` / `.xlsm` 会被明确拒绝，不会静默转换
- 仅 Windows；WPS 定位依赖本机 WPS 的 COM 注册情况
- Windows 10 真机与 WPS 的完整验收尚未执行，见 [docs/TEST_REPORT.md](docs/TEST_REPORT.md)
- 1.x 的数据目录（`PartNumberSearch`）与 2.0 的数据库结构不通用，升级后需要重新导入工作簿

## 许可

[MIT](LICENSE) © 2026 Forken

`scripts/ChineseSimplified.isl` 来自 [Inno Setup 官方仓库](https://github.com/jrsoftware/issrc/blob/main/Files/Languages/ChineseSimplified.isl)
的简体中文翻译，保留原作者署名，未作修改。
