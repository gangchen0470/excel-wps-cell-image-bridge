from __future__ import annotations

from dataclasses import dataclass, asdict


@dataclass(frozen=True)
class CellImage:
    sheet_name: str
    sheet_path: str
    cell: str
    media_path: str
    source_format: str
    image_id: str
    content_type: str | None = None

    def as_dict(self) -> dict[str, str | None]:
        return asdict(self)


@dataclass(frozen=True)
class FloatingImage(CellImage):
    drawing_path: str = ""
    relationship_id: str = ""
    anchor_index: int = 0


class InvalidWorkbookError(ValueError):
    pass


class UnsupportedWorkbookError(ValueError):
    pass
