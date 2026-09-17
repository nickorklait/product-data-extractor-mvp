import os
import re
from datetime import datetime, timezone
from pathlib import Path


VOLUME_PATH_ENV = "PRODUCT_DATA_VOLUME_PATH"


class VolumeArchive:
    def __init__(self, root: Path | None = None) -> None:
        self.root = root

    @classmethod
    def from_environment(cls) -> "VolumeArchive":
        configured_path = os.getenv(VOLUME_PATH_ENV, "").strip()
        return cls(Path(configured_path) if configured_path else None)

    @property
    def enabled(self) -> bool:
        return self.root is not None

    def write(
        self,
        category: str,
        filename: str,
        content: bytes,
        processing_id: str,
        created_at: datetime | None = None,
    ) -> Path | None:
        if not self.root:
            return None
        if category not in {"source", "failed", "excel"}:
            raise ValueError(f"Unsupported archive category: {category}")

        timestamp = created_at or datetime.now(timezone.utc)
        target_dir = self.root / category / timestamp.strftime("%Y") / timestamp.strftime("%m") / timestamp.strftime("%d")
        target_dir.mkdir(parents=True, exist_ok=True)

        safe_name = _safe_filename(filename)
        safe_processing_id = _safe_identifier(processing_id)
        target = target_dir / f"{safe_processing_id}_{safe_name}"
        target.write_bytes(content)
        return target


def _safe_filename(filename: str) -> str:
    source_name = Path(filename).name.strip()
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", source_name).strip("._")
    return safe_name[:180] or "document"


def _safe_identifier(value: str) -> str:
    safe_value = re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_-")
    return safe_value[:64] or "unknown"
