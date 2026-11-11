# 测试报告

## 环境

| 项目 | 值 |
| --- | --- |
| 日期 | 2026-11-11 |
| 系统 | Windows 11 x64（开发机，**不构成** Windows 10 验收） |
| Python | 3.12.10 (CPython AMD64) |
| PySide6 | 6.12.0 |
| openpyxl | 3.1.5 |
| pytest / pytest-qt | 9.2.0 / 4.5.0 |
| Qt 平台 | `offscreen`（无显示器也能跑） |

## 命令

```powershell
.venv\Scripts\python.exe -m pytest -q
.venv\Scripts\python.exe -m app.main --self-test
```

## 结果

```
72 passed in 3.99s
```

`--self-test` 通过：在临时目录用合成工作簿跑真实的 Parser、SQLite 和界面（offscreen），
并验证重复文件会被拒绝。

## 覆盖面

| 文件 | 用例数 | 覆盖内容 |
| --- | --- | --- |
| `tests/unit/test_excel_parser.py` | 29 | 表头识别、多区域、合并单元格、重复表头、公式缓存、空 fill 修复、损坏文件、非法输入 |
| `tests/unit/test_excel_parser_regressions.py` | 4 | 历史回归：文本行保留、共享公式展开 |
| `tests/unit/test_inventory_headers.py` | 2 | 库存类表头识别 |
| `tests/unit/test_import_service.py` | 10 | 添加/替换/移除、6 个以上工作簿、重复 SHA-256、失败保留旧索引、取消、进度回调、相对路径与格式拒绝 |
| `tests/unit/test_search_service.py` | 11 | 逐单元格包含、大小写、首尾空格、跨单元格不拼接、单行多命中去重、跨工作簿顺序、分页、通配符按字面处理、参数校验、完整行格式化 |
| `tests/unit/test_contracts.py` | 8 | DTO 不可变、服务实现 Protocol、接口无固定槽位 API、错误码 |
| `tests/ui/test_pages.py` | 8 | 数据源列表与空状态、检索结果渲染、零结果、未索引时引导、页面切换、底部计数、导航互斥 |

测试全部使用 `tests/support.py` 现场生成的合成工作簿，不依赖任何真实业务文件。

## 未验证项

这些必须在目标环境验证后才可以宣称"已验收"：

| 项目 | 状态 |
| --- | --- |
| PyInstaller 打包 + 打包后自检 | **已验证**：on-dir 产物 118 MB、187 个文件，`MultiTableSearch.exe --self-test` 退出码 0 |
| Windows 10 1809+ 真机运行 | 未执行（开发机是 Windows 11） |
| WPS COM 打开并定位到单元格 | 未执行（当前用 `RecordingOpener` 替身；`app/windows/wps_opener.py` 有真实实现但未联调） |
| 真实大文件导入性能（20MB+、28 个工作表） | 只做过人工观察，未做目标环境基准 |
| Inno Setup 安装包与卸载 | 未执行（构建机未安装 Inno Setup 6） |
| 100% / 150% / 200% DPI | 未执行 |
| 长时间运行与并发取消 | 未执行 |

## 已知限制

- 只支持 `.xlsx`；`.xls` / `.xlsm` 明确拒绝
- 1.x 的数据库（schema v1）与 2.0 不兼容，数据目录也从 `PartNumberSearch` 改为
  `MultiTableSearch`，升级后需要重新导入
- 文件名去重策略是内容 SHA-256：同内容换个文件名仍然会被拒绝
