from __future__ import annotations
from typing import Protocol


class TermProto(Protocol):
    """A protocol for version objects."""

    def __init__(self) -> None:
        """Initialize the version object."""
        ...

    def get_is_match(self) -> bool:
        """Check if the terminal is a match"""
        ...

    def start(self, msg: str, title: str = "Terminal") -> None:
        """Start the terminal."""
        ...

    def stop(self) -> None:
        """Stop the terminal."""
        ...
