"""Base class for everything a client sends: unknown fields are errors, not silently ignored."""

from typing import Any

from pydantic import BaseModel, ConfigDict, model_validator
from pydantic_core import PydanticCustomError

from excel_mcp.errors import InvalidArgumentError


class InputModel(BaseModel):
    model_config = ConfigDict(extra="forbid")

    @model_validator(mode="before")
    @classmethod
    def _reject_unknown_fields(cls, data: Any) -> Any:
        if isinstance(data, dict):
            unknown = [str(key) for key in data if key not in cls.model_fields]
            if unknown:
                raise PydanticCustomError(
                    "unknown_fields",
                    "unknown fields",
                    {"unknown": unknown, "valid": list(cls.model_fields)},
                )
        return data

    def reject_unused(self, used: set[str], kind: str) -> None:
        """Fail when a field that does not apply to ``kind`` was set."""
        unused = self.model_fields_set - used
        if unused:
            raise InvalidArgumentError(
                f"{kind} does not use {', '.join(sorted(unused))}. "
                f"It takes: {', '.join(sorted(used))}."
            )
