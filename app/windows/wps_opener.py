"""Registered WPS COM automation with truthful default-association fallback."""
import logging
import os
from pathlib import Path
import re
import threading

from app.contracts import ContractError, ErrorCode, OpenResult, OpenStatus

log = logging.getLogger(__name__)


def registered_wps_progids() -> tuple[str, ...]:
    if os.name != "nt":
        return ()
    import winreg
    found = set()
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "", 0, winreg.KEY_READ | view) as root:
                index = 0
                while True:
                    try:
                        name = winreg.EnumKey(root, index)
                    except OSError:
                        break
                    index += 1
                    if re.fullmatch(r"(?:ket|et)\.Application(?:\.\d+)?", name, re.I):
                        found.add(name)
        except OSError:
            continue
    return tuple(sorted(found, key=lambda x: (not x.lower().startswith("ket."), x)))


class WindowsWpsOpener:
    def __init__(self):
        self._lock = threading.Lock()

    def open_location(self, path: str, sheet_name: str, cell_address: str) -> OpenResult:
        match = re.fullmatch(r"([A-Z]{1,3})([1-9]\d*)", cell_address) if isinstance(cell_address, str) else None
        col = 0
        if match:
            for letter in match[1]:
                col = col * 26 + ord(letter) - 64
        if (not isinstance(path, str) or not Path(path).is_absolute()
                or not isinstance(sheet_name, str) or not sheet_name or not match
                or col > 16384 or int(match[2]) > 1048576):
            raise ContractError(ErrorCode.INVALID_ARGUMENT, "文件绝对路径、工作表或 A1 坐标不合法")
        source = Path(path)
        if not source.is_file():
            return OpenResult(OpenStatus.FAILED, path, sheet_name, cell_address,
                              "原始副本不存在，请恢复数据备份或重新导入", ErrorCode.FILE_NOT_FOUND)
        with self._lock:
            try:
                self._locate(source, sheet_name, cell_address, int(match[2]), col)
                return OpenResult(OpenStatus.LOCATED, path, sheet_name, cell_address, "已在 WPS 中定位命中单元格")
            except Exception:
                log.exception("WPS 自动定位失败，尝试默认关联打开")
                error = ErrorCode.WPS_UNAVAILABLE if not registered_wps_progids() else ErrorCode.OPEN_FAILED
            try:
                os.startfile(str(source))
                return OpenResult(OpenStatus.OPENED_WITHOUT_LOCATION, path, sheet_name, cell_address,
                                  f"已请求默认程序打开，未能自动定位。请手动选择工作表“{sheet_name}”，跳转 {cell_address}。", error)
            except (OSError, AttributeError):
                log.exception("默认关联打开失败")
                return OpenResult(OpenStatus.FAILED, path, sheet_name, cell_address,
                                  f"WPS 自动化及默认关联均无法打开；可手动打开文件并选择“{sheet_name}” {cell_address}", ErrorCode.OPEN_FAILED)

    def _locate(self, source, sheet_name, address, row, column):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = book = sheet = target = selection = None
        try:
            last_error = None
            for progid in registered_wps_progids():
                try:
                    try:
                        candidate = win32com.client.GetActiveObject(progid)
                    except Exception:
                        candidate = win32com.client.Dispatch(progid)
                    candidate.Workbooks.Count
                    app = candidate
                    break
                except Exception as exc:
                    last_error = exc
            if app is None:
                raise RuntimeError("未找到可用 WPS 表格 COM 接口") from last_error
            app.Visible = True
            wanted = os.path.normcase(os.path.abspath(source))
            for index in range(1, app.Workbooks.Count + 1):
                item = app.Workbooks.Item(index)
                if os.path.normcase(os.path.abspath(str(item.FullName))) == wanted:
                    book = item
                    break
            if book is None:
                old_security = app.AutomationSecurity
                try:
                    app.AutomationSecurity = 3
                    book = app.Workbooks.Open(str(source), 0, True)
                finally:
                    app.AutomationSecurity = old_security
            book.Activate()
            sheet = book.Worksheets.Item(sheet_name)
            sheet.Activate()
            target = sheet.Range(address)
            target.Select()
            selection = app.Selection
            if (str(app.ActiveSheet.Name) != sheet_name
                    or os.path.normcase(os.path.abspath(str(app.ActiveWorkbook.FullName))) != wanted
                    or not (selection.Row <= row < selection.Row + selection.Rows.Count
                            and selection.Column <= column < selection.Column + selection.Columns.Count)):
                raise RuntimeError("WPS 未确认已选中请求的工作表与单元格")
            try:
                import win32gui
                win32gui.ShowWindow(int(app.Hwnd), 9)
                win32gui.SetForegroundWindow(int(app.Hwnd))
            except Exception:
                log.info("已定位，但 Windows 未允许切换到前台")
        finally:
            selection = target = sheet = book = app = None
            pythoncom.CoUninitialize()
