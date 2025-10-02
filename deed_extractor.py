"""Utility helpers for cleaning deed text before parsing bearings and distances."""

from __future__ import annotations

import re
from typing import Iterable

_CHAR_NORMALIZE_MAP = {
    "′": "'",
    "’": "'",
    "`": "'",
    "‛": "'",
    "＇": "'",
    "¨": "'",
    "˝": '"',
    "″": '"',
    "“": '"',
    "”": '"',
    "„": '"',
    "º": "°",
    "˚": "°",
    "⁰": "°",
    "°": "°",
    "‐": "-",
    "‑": "-",
    "‒": "-",
    "–": "-",
    "—": "-",
    "―": "-",
}

_DEGREE_WORD_PATTERN = re.compile(r"(?i)\bDEG(?:REE|REES)?\b")
_MINUTE_WORD_PATTERN = re.compile(r"(?i)\bMIN(?:UTE|UTES)?\b")
_SECOND_WORD_PATTERN = re.compile(r"(?i)\bSEC(?:OND|ONDS)?\b")
_CARDINAL_PATTERN = re.compile(
    r"(?i)\b(NORTH|SOUTH|EAST|WEST)(?:\s+|-)?(EAST|WEST)?(?:ERLY)?\b"
)
_CARDINAL_ABBREV_PATTERN = re.compile(r"(?i)\b([NSEW])\.(?=\s|$)")
_LETTER_O_DEGREE_PATTERN = re.compile(r"(?<=\d)\s*[oO](?=\s*\d)")
_MULTI_SPACE_PATTERN = re.compile(r"\s+")
_CUE_WORD_PATTERN = re.compile(
    r"(?i)\b(THENCE|THEN|BEGINNING|BEGIN|ENDING|CONTAINING|WITH|ALONG|RUNNING)\b"
)
_PUNCT_GAP_PATTERN = re.compile(r"[;,]*(?=\s)")
_TRAILING_DIR_PUNCT_PATTERN = re.compile(r"(?i)\b([NSEW]{1,2})[\.,;:]+(?=\s)")
_UNIT_PUNCT_PATTERN = re.compile(
    r"(?i)\b(FEET|FT|FOOT|METERS|M|CHAINS|CHS|CHAIN|RODS|RDS|ROD)[\.,;:]+(?=\s)"
)
_HEADER_FOOTER_PATTERNS: Iterable[re.Pattern[str]] = (
    re.compile(r"(?i)^\s*page\s+\d+(?:\s+of\s+\d+)?\s*$"),
    re.compile(r"(?i)^\s*continued\s*$"),
    re.compile(r"(?i)^\s*-{2,}\s*$"),
)


def _strip_headers_and_footers(text: str) -> str:
    lines = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if any(pattern.match(line) for pattern in _HEADER_FOOTER_PATTERNS):
            continue
        lines.append(line)
    return " ".join(lines)


def _normalize_special_chars(text: str) -> str:
    if not text:
        return ""
    chars = []
    for ch in text:
        chars.append(_CHAR_NORMALIZE_MAP.get(ch, ch))
    normalized = "".join(chars)
    normalized = _DEGREE_WORD_PATTERN.sub("°", normalized)
    normalized = _LETTER_O_DEGREE_PATTERN.sub("°", normalized)
    normalized = _MINUTE_WORD_PATTERN.sub("'", normalized)
    normalized = _SECOND_WORD_PATTERN.sub('"', normalized)
    normalized = normalized.replace("º", "°")
    return normalized


def _standardize_cardinals(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        primary = match.group(1)[0].upper()
        secondary = match.group(2)
        if secondary:
            return primary + secondary[0].upper()
        return primary

    text = _CARDINAL_PATTERN.sub(repl, text)
    text = _CARDINAL_ABBREV_PATTERN.sub(lambda m: m.group(1).upper(), text)
    return text


def _uppercase_cues(text: str) -> str:
    return _CUE_WORD_PATTERN.sub(lambda m: m.group(1).upper(), text)


def clean_text(raw: str) -> str:
    """Normalize deed text into a compact, parser-friendly representation."""

    if not raw:
        return ""

    text = raw.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("-\n", "")
    text = _strip_headers_and_footers(text)
    text = _normalize_special_chars(text)
    text = _standardize_cardinals(text)
    text = _uppercase_cues(text)
    text = _TRAILING_DIR_PUNCT_PATTERN.sub(lambda m: m.group(1).upper(), text)
    text = _UNIT_PUNCT_PATTERN.sub(lambda m: m.group(1).lower(), text)
    text = _PUNCT_GAP_PATTERN.sub("", text)
    text = re.sub(r"(?<=\d),(?=\d)", ",", text)  # keep decimals/commas in numbers
    text = re.sub(r"(?<=\d)[,](?=\s)", "", text)
    text = re.sub(r"(?<=\b)([NSEW])(?=\s+\d)", lambda m: m.group(1).upper(), text)
    text = _MULTI_SPACE_PATTERN.sub(" ", text)
    return text.strip()


__all__ = ["clean_text"]

