"""Base class for everything a client sends: unknown fields are errors, not silently ignored."""

from typing import Any

from pydantic import BaseModel, ConfigDict, model_validator
from pydantic_core import PydanticCustomError


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
