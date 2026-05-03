from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

__all__ = ["RepeatRecords", "repeat_records", "is_repeat_records"]

# §1. Constants & Exceptions

# §2. Classes and Sub Classes


@dataclass(frozen=True)
class RepeatRecords:
    """Source-backed repeat records with shared per-record constants."""

    records: Any
    constants: dict[str, Any]


# §3. Private Helper Functions

# §4. Public Functions


def repeat_records(
    records: Any | Iterable[dict[str, Any]],
    constants: dict[str, Any] | None = None,
) -> RepeatRecords:
    return RepeatRecords(records=records, constants=dict(constants or {}))


def is_repeat_records(value: Any) -> bool:
    return isinstance(value, RepeatRecords)


# §5. Entrypoints
