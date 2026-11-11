# 架构

目标：把"任意多个工作簿 → 一个本地索引 → 逐单元格检索"这条链路拆成可独立测试的模块，
并保证任何失败都不会破坏已索引的数据。

## 分层

```
app/main.py                 组合根：拼装服务、解析命令行、安装自检
        │
app/ui/                     Qt 界面（只依赖 contracts，不直接碰 openpyxl / sqlite3 / COM）
        │
app/contracts.py            唯一接口层：DTO + Protocol，标准库实现，无第三方依赖
        │
app/core/ + app/excel_parser/ + app/windows/
```

依赖方向永远向下。`app/contracts.py` 不 import Qt、openpyxl、sqlite3 或 pywin32，
因此解析器可以在无界面、无数据库的环境里单独测试。

## 接口

`app/contracts.py` 的 `INTERFACE_VERSION = "2.0"`。四个服务：

| Protocol | 职责 |
| --- | --- |
| `DataManager` | `list_workbooks`、`add_workbook`、`replace_workbook`、`remove_workbook` |
| `SearchService` | `search(keyword, offset, limit) -> SearchPage` |
| `WpsOpener` | `open_location(path, sheet, cell) -> OpenResult` |
| `ResultFormatter` | `format_results(results) -> str` |

关键 DTO：`ParsedWorkbook`/`ParsedSheet`/`ParsedTable`/`ParsedRow`/`ParsedCell`（解析产物）、
`WorkbookInfo`（已索引工作簿）、`SearchResult`（一条命中）、`ImportSummary`（一次导入结果）。
全部是 `frozen=True, slots=True` 的不可变对象，集合一律用 tuple。

**没有固定槽位。** 工作簿用数据库分配的 `workbook_id: int` 标识，`WorkbookInfo` 是可枚举
列表里的一项，列表长度没有上限。

错误统一抛 `ContractError(code, message, details)`；`message` 是可以直接显示给用户的中文。

### 参数错误 vs 运行失败

- 调用方传了非法参数（相对路径、`workbook_id <= 0`）→ **抛** `ContractError(INVALID_ARGUMENT)`
- 进入流水线后的失败（文件损坏、磁盘错误、取消）→ **返回** `ImportSummary(success=False, error_code=...)`

这条区分让 UI 可以只在外层 catch 一次参数错误，把流水线失败渲染成卡片状态。

## 线程模型

Qt 界面线程只做三件事：收集输入、发起后台任务、渲染结果。

`app/ui/workers.py` 的 `TaskRunner` 把阻塞调用丢进全局 `QThreadPool`，用 queued signal
把结果送回界面线程。导入进度回调在 worker 线程触发，只允许 `emit`，不能碰 QWidget。
检索用递增的 `_request_id` 丢弃过期返回，避免慢请求覆盖新关键词的结果。

## 导入状态机

```
复制到 staging ─→ SHA-256 双向校验 ─→ 解析 ─→ 行/坐标/区域校验 ─→ 备份 DB ─→ 单事务提交
      │                  │              │            │                │           │
      └──────────────────┴──────────────┴────────────┴────────────────┴──→ 失败：回滚
```

- 每一步之前检查取消令牌；**提交阶段不可取消**
- 托管副本先落盘（`os.replace` + `chmod 0o444`），数据库行后才引用它
- 写库是 `BEGIN IMMEDIATE` 的单事务：要么全部生效，要么完全回滚
- 失败时保留旧 `workbooks` 行、旧工作表数据和旧托管副本；只把状态标记为 `failed`
- 同一个文件（SHA-256 相同）不允许重复索引
- 替换成功后才删除被替换的托管副本目录（Windows 下要先清只读位）

## 数据模型

SQLite，WAL 模式，`schema_version = 2`。

```
workbooks(workbook_id PK AUTOINCREMENT, original_filename, source_path, managed_path,
          sha256 UNIQUE, file_size_bytes, imported_at, sheet_count, row_count, import_version)
worksheets(id PK, workbook_id FK→workbooks ON DELETE CASCADE, sheet_name, sheet_order)
tables(id PK, sheet_id FK, table_index, table_name, min_row/max_row/min_col/max_col,
       header_rows, merged_ranges)
rows(id PK, sheet_id FK, row_number, cells_json, table_refs)
row_table_memberships(row_id, table_id)
cells(row_id FK, table_id FK, column_index, column_letter, header, raw_json, display_value,
      address, is_supplementary, data_type, number_format, origin_address,
      formula_has_cached_value, search_text)
import_history(id PK, workbook_id, action, version, sha256, created_at, success, message, backup_path)
```

`cells.search_text` 是 `display_value.strip().casefold()` 的物化副本；检索只对它做
`instr()`，因此匹配语义等于 Python 的 `in`，不需要 LIKE 转义。

级联删除意味着移除一个工作簿只需要 `DELETE FROM workbooks WHERE workbook_id=?`。

## 检索语义

1. 关键词 `strip()` 后 `casefold()`
2. 在**单元格**级别做 `instr(search_text, ?) > 0`，绝不拼接相邻单元格
3. 命中行按 `workbook_id, sheet_order, row_number` 稳定排序并去重
4. 同一行多个单元格命中只输出一次，`matched_cell_addresses` 保留全部位置
5. 分页在 SQL 之外做切片，`total_count` 是完整命中数
6. 命中结果分页时按 `(keyword, [(workbook_id, import_version)])` 缓存行 id 列表；
   任一次导入都会改变这个 key，缓存自然失效

## 界面

单向依赖：`ui/*` 依赖 `contracts`，`main.py` 依赖一切。
`create_main_window(services)` 只是工厂，不创建 `QApplication`、不进入事件循环，
因此测试可以在 `offscreen` 平台直接构造窗口。

设计系统与色板见 [UI_DESIGN.md](UI_DESIGN.md)。

## 扩展点

- 支持 `.xlsm`：在 `parse_workbook` 前后放开后缀判断即可，数据库无需改动
- 文件夹监控导入：新增一个 `DataManager` 调用方，`add_workbook` 已是幂等的（重复 SHA-256 会被拒绝）
- 结果导出：实现另一个 `ResultFormatter`，界面只依赖 Protocol
