# 参与开发

欢迎 issue 和 pull request。这个项目的目标很小：**把料号查得又准又快，并且绝不改动原始数据**。
任何让结果"看起来更好"但可能改变数据的改动都不会被接受。

## 环境

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest -q
```

## 约定

- Python 3.12，Windows 目标；不要引入需要联网的运行时依赖
- 模块之间只通过 `app/contracts.py` 交互，不要在 UI 里直接调用 openpyxl、sqlite3 或 COM
- 阻塞调用放在 `app/ui/workers.py` 的后台任务里，界面线程只处理 Qt 对象
- 数据改动前必须有备份，提交必须是单个事务；失败要保留旧索引
- 新增或修改行为请带上测试。测试使用合成工作簿，**不要提交任何真实业务文件**
- 提交前跑一遍 `pytest`，并检查 `git status` 里没有 `.xlsx` / `.db` / 日志

## 测试数据

`tests/support.py` 里的 `build_workbook()` 会在临时目录生成合成工作簿，表头用
`日期/型号/供应商/备注` 这类通用名字。请沿用虚构的供应商名（示例供应商A、供应商甲……），
不要在仓库里出现真实的客户、供应商或料号。

## 提交信息

用普通的一句话说明改了什么，例如 `fix: 导入失败时保留旧索引`。
不需要签名或模板。
