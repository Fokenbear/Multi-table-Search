"""Background execution for blocking service calls.

Per docs/ARCHITECTURE.md the injected services are synchronous and blocking:
every call runs on the global QThreadPool and results/errors return to the UI
thread through queued Qt signals.  Service kwargs such as `on_progress` and
`cancellation` are passed through untouched — hook keywords here are
`finished` / `failed` so they never collide with service parameters.
Progress callbacks fire on the worker thread and must only emit Qt signals.
"""

from __future__ import annotations

import traceback
from threading import Event
from typing import Any, Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot

from app.contracts import CancellationToken, ContractError, ErrorCode


class CancellableToken:
    """Concrete CancellationToken for import cancellation requests."""

    def __init__(self) -> None:
        self._cancelled = Event()

    def cancel(self) -> None:
        self._cancelled.set()

    def is_cancelled(self) -> bool:
        return self._cancelled.is_set()


class _TaskSignals(QObject):
    finished = Signal(object)
    failed = Signal(object)  # ContractError


class ServiceTask(QRunnable):
    """Run one blocking callable; publish result / ContractError via signals."""

    def __init__(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> None:
        super().__init__()
        self._fn = fn
        self._args = args
        self._kwargs = kwargs
        self.signals = _TaskSignals()
        self.setAutoDelete(True)

    @Slot()
    def run(self) -> None:
        try:
            result = self._fn(*self._args, **self._kwargs)
        except ContractError as exc:
            self.signals.failed.emit(exc)
        except Exception as exc:  # noqa: BLE001 - worker must not crash the app
            wrapped = ContractError(
                ErrorCode.INTERNAL_ERROR,
                f"发生内部错误：{exc}",
                details=traceback.format_exc(),
            )
            self.signals.failed.emit(wrapped)
        else:
            self.signals.finished.emit(result)


class TaskRunner(QObject):
    """Owns live tasks so their signal objects survive until delivery."""

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._tasks: list[ServiceTask] = []

    @property
    def active_count(self) -> int:
        return len(self._tasks)

    def submit(
        self,
        fn: Callable[..., Any],
        *args: Any,
        finished: Callable[[Any], None] | None = None,
        failed: Callable[[ContractError], None] | None = None,
        **kwargs: Any,
    ) -> ServiceTask:
        task = ServiceTask(fn, *args, **kwargs)
        if finished is not None:
            task.signals.finished.connect(finished)
        if failed is not None:
            task.signals.failed.connect(failed)
        task.signals.finished.connect(lambda _r, t=task: self._release(t))
        task.signals.failed.connect(lambda _e, t=task: self._release(t))
        self._tasks.append(task)
        QThreadPool.globalInstance().start(task)
        return task

    def _release(self, task: ServiceTask) -> None:
        try:
            self._tasks.remove(task)
        except ValueError:
            pass


def signal_emitter(signal: Signal) -> Callable[[Any], None]:
    """Adapt a DataManager.on_progress callback to a queued Qt signal."""

    def _forward(payload: Any) -> None:
        signal.emit(payload)

    return _forward


def as_cancellation(token: CancellableToken | None) -> CancellationToken | None:
    return token
