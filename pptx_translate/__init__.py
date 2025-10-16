"""
Helper utilities and translators for PPTX translation CLI.

This package exposes helpers for collecting text items, applying glossary rules,
and selecting different OpenAI API backends.
"""

from .glossary import apply_glossary, read_glossary
from .pptx_utils import (
    TextItem,
    collect_items,
    get_text,
    iter_text_frames,
    parse_slide_spec,
    set_text,
    summarize_deck,
)

__all__ = [
    "TextItem",
    "collect_items",
    "iter_text_frames",
    "get_text",
    "set_text",
    "summarize_deck",
    "parse_slide_spec",
    "read_glossary",
    "apply_glossary",
]
