"""User data directories, independent of the executable working directory."""
from dataclasses import dataclass
import os
from pathlib import Path


@dataclass(frozen=True)
class AppPaths:
    root: Path

    @classmethod
    def create(cls, root: str | Path | None = None):
        if root is None:
            base = os.environ.get("LOCALAPPDATA")
            if not base:
                raise RuntimeError("无法确定 LOCALAPPDATA 用户数据目录")
            root = Path(base) / "MultiTableSearch"
        result = cls(Path(root).resolve())
        for folder in ("database", "workbooks", "backups", "logs", "settings", "staging"):
            (result.root / folder).mkdir(parents=True, exist_ok=True)
        return result

    @property
    def database(self) -> Path:
        return self.root / "database" / "search.db"
