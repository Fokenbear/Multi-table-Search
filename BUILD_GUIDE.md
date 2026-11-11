# 构建指南

目标：在 Windows 上产出一个自带 Python 运行时、双击即用的安装包。

- 目标系统：Windows 10 1809（build 17763）+ x64
- 构建机：Windows 10/11 x64，Python 3.12
- 应用版本：2.0.0

## 1. 准备

在仓库根目录：

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pip check
```

开发依赖包含 PyInstaller 6.22.3。打包进成品的运行时依赖见 `requirements.txt`；
构建机实测的完整版本清单见 `requirements-environment.txt`。

安装 [Inno Setup 6](https://jrsoftware.org/isdl.php)（只需要做安装包时用）。

## 2. 先测试，再构建

```powershell
.venv\Scripts\python.exe -m pytest -q
.venv\Scripts\python.exe -m app.main --self-test
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

`build_windows.ps1` 依次做：

1. `pip check`
2. PyInstaller `--onedir --windowed`，产物在 `build\bundle-<时间戳>\MultiTableSearch\`
3. 检查产物里没有 `xlsx` / `db` / `csv` 等数据文件，有就中止
4. 把本次产物路径写进 `build\last-bundle-path.txt`
5. 调用 ISCC 生成 `dist\MultiTableSearch_Setup.exe` 并打印 SHA-256

只想验证打包、不做安装包时加 `-SkipInstaller`。

入口是 `app\launcher.py`（PyInstaller 需要绝对导入的入口文件），GUI 用
`--windowed`，控制台不弹出。`tests`、`scripts`、pytest 不会进入产物。

## 3. 安装包

`scripts/create_installer.iss` 是 Inno Setup 6 脚本，由打包脚本用
`/DBundleDir=<产物目录>` 调用。

- 默认安装到 `%LOCALAPPDATA%\Programs\MultiTableSearch`，不需要管理员权限
- 可选创建桌面快捷方式
- 简体中文语言文件 `scripts/ChineseSimplified.isl` 来自 Inno Setup 官方仓库
  （<https://github.com/jrsoftware/issrc/blob/main/Files/Languages/ChineseSimplified.isl>），
  下载来源见 <https://jrsoftware.org/files/istrans/>，保留原文件署名
- 卸载不删除用户数据目录

Qt 6.12 要求 Windows 10 1809 或更新，因此 `MinVersion=10.0.17763`，与
[Qt 官方平台说明](https://doc.qt.io/qt-6.12/windows.html) 一致。这不能代替真机实测。

## 4. 安装后自检

```powershell
"%LOCALAPPDATA%\Programs\MultiTableSearch\MultiTableSearch.exe" --self-test
```

自检用合成工作簿在临时目录跑真实的解析、SQLite 和界面（offscreen），
会打印一行 JSON；退出码为 0 表示通过。

## 5. 发布前检查

- [ ] `pytest` 全绿
- [ ] `--self-test` 通过
- [ ] 产物目录里没有业务文件、数据库、日志
- [ ] 在**目标系统**（Windows 10 x64 + WPS）上安装、导入、检索、定位、卸载各跑一遍
- [ ] 100% / 150% / 200% DPI 下界面不截断
- [ ] 记录安装包 SHA-256

当前哪些项已经做过、哪些没做，见 [docs/TEST_REPORT.md](docs/TEST_REPORT.md)。
