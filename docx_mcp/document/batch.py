"""Batch mixin: apply several edit operations in one call, all or nothing."""

from __future__ import annotations

import contextlib
import copy
import inspect
from typing import Any

_BATCH_OPS = (
    "replace_text",
    "insert_text",
    "delete_text",
    "add_comment",
    "modify_cell",
    "insert_paragraph",
    "update_paragraph",
    "delete_paragraph",
)
_MAX_OPS = 500


class BatchMixin:
    """Atomic multi-operation editing."""

    def batch_edit(self, operations: list[dict], *, author: str = "Claude") -> dict:
        """Apply *operations* in order. If any fails, none of them take effect.

        Each operation is ``{"op": <name>, <parameters of that edit>}`` where the name is one
        of replace_text, insert_text, delete_text, add_comment, modify_cell, insert_paragraph,
        update_paragraph, delete_paragraph. Parameters are exactly those of the matching tool.
        *author* is used for every operation that takes one and does not set its own.
        """
        calls = self._prepare_batch(operations, author)
        snapshot = self._snapshot_state()
        results: list[dict[str, Any]] = []
        for index, (name, kwargs) in enumerate(calls):
            try:
                result = getattr(self, name)(**kwargs)
            except Exception as exc:
                self._restore_state(snapshot)
                raise ValueError(
                    f"Operation {index} ({name}) failed: {exc}. No changes were applied."
                ) from exc
            results.append({"index": index, "op": name, "result": result})
        return {"applied": True, "count": len(results), "results": results}

    def _prepare_batch(self, operations: list[dict], author: str) -> list[tuple[str, dict]]:
        """Validate every operation before any is applied; return ``(method, kwargs)`` pairs."""
        if not isinstance(operations, list) or not operations:
            raise ValueError("operations must be a non-empty list of objects")
        if len(operations) > _MAX_OPS:
            raise ValueError(f"Too many operations: {len(operations)} (limit {_MAX_OPS})")

        calls: list[tuple[str, dict]] = []
        for index, entry in enumerate(operations):
            if not isinstance(entry, dict):
                raise ValueError(f"Operation {index} must be an object, got {type(entry).__name__}")
            name = entry.get("op")
            if name not in _BATCH_OPS:
                raise ValueError(
                    f"Operation {index}: unknown op {name!r}; "
                    f"expected one of {', '.join(_BATCH_OPS)}"
                )
            params = inspect.signature(getattr(self, name)).parameters
            kwargs = {k: v for k, v in entry.items() if k != "op"}
            unknown = sorted(set(kwargs) - set(params))
            if unknown:
                raise ValueError(
                    f"Operation {index} ({name}): unknown parameter(s) {', '.join(unknown)}; "
                    f"allowed: {', '.join(params)}"
                )
            if "author" in params:
                kwargs.setdefault("author", author)
            missing = [
                p for p, spec in params.items() if spec.default is spec.empty and p not in kwargs
            ]
            if missing:
                raise ValueError(f"Operation {index} ({name}): missing {', '.join(missing)}")
            for key, value in kwargs.items():
                # JSON numbers can arrive as 3.0; index parameters must be real ints.
                if (
                    params[key].annotation == "int"
                    and isinstance(value, float)
                    and value.is_integer()
                ):
                    kwargs[key] = int(value)
            calls.append((name, kwargs))
        return calls

    def _snapshot_state(self) -> dict[str, Any]:
        return {
            "trees": {name: copy.deepcopy(tree) for name, tree in self._trees.items()},
            "modified": set(self._modified),
            "binaries": dict(self._binaries),
            "files": self._workdir_files(),
        }

    def _restore_state(self, snapshot: dict[str, Any]) -> None:
        self._trees.clear()
        self._trees.update(snapshot["trees"])
        self._modified = snapshot["modified"]
        self._binaries = snapshot["binaries"]
        # save() repacks everything in the work dir, so drop parts an operation created.
        for path in self._workdir_files() - snapshot["files"]:
            with contextlib.suppress(OSError):
                path.unlink()

    def _workdir_files(self) -> set:
        return {p for p in self.workdir.rglob("*") if p.is_file()} if self.workdir else set()
