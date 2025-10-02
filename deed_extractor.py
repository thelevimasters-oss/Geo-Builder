"""Utility helpers for extracting deed information from raw text."""

from __future__ import annotations

import re
from typing import Iterable

# Character replacements that normalize various unicode punctuation into
# straightforward ASCII equivalents that downstream parsing expects.
_CHAR_NORMALIZE_MAP = {
    "\u2018": "'",
    "\u2019": "'",
    "\u201a": "'",
    "\u2032": "'",
    "\u2035": "'",
    "\u201c": '"',
    "\u201d": '"',
    "\u2033": '"',
    "\u2036": '"',
    "\u00ba": "°",
    "\u00b0": "°",
    "\u2010": "-",
    "\u2013": "-",
    "\u2014": "-",
    "\u2212": "-",
    "\u00a0": " ",
}

_COMPOUND_CARDINAL_PATTERN = re.compile(
    r"(?i)\b(NORTH|SOUTH)(?:\s+|-)?(EAST|WEST)(?:ERLY)?\b"
)
_CARDINAL_WORD_MAP = {
    "NORTH": "N",
    "NORTHERLY": "N",
    "SOUTH": "S",
    "SOUTHERLY": "S",
    "EAST": "E",
    "EASTERLY": "E",
    "WEST": "W",
    "WESTERLY": "W",
}

_HEADER_FOOTER_PATTERNS: Iterable[re.Pattern[str]] = (
    re.compile(r"(?i)^page\s+\d+(?:\s*(?:of|/)\s*\d+)?$"),
    re.compile(r"(?i)^-+\s*page\s*\d+\s*-+$"),
    re.compile(r"(?i)^\d+\s*/\s*\d+$"),
)

_CUE_WORDS = (
    "THENCE",
    "BEGINNING",
    "RUNNING",
    "CONTINUING",
    "CONTAINING",
    "ALONG",
)


def _normalize_characters(text: str) -> str:
    for original, replacement in _CHAR_NORMALIZE_MAP.items():
        text = text.replace(original, replacement)
    return text


def _remove_headers_and_footers(text: str) -> str:
    lines = text.splitlines()
    cleaned_lines = []
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if any(pattern.match(line) for pattern in _HEADER_FOOTER_PATTERNS):
            continue
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def _standardize_cardinals(text: str) -> str:
    def _compound_repl(match: re.Match[str]) -> str:
        primary = match.group(1)[0].upper()
        secondary = match.group(2)[0].upper()
        suffix = "ERLY" if match.group(0).rstrip().upper().endswith("ERLY") else ""
        return f"{primary}{secondary}{suffix}"

    text = _COMPOUND_CARDINAL_PATTERN.sub(_compound_repl, text)

    def _replace_single(match: re.Match[str]) -> str:
        return _CARDINAL_WORD_MAP[match.group(0).upper()]

    pattern = re.compile(r"(?i)\b(" + "|".join(_CARDINAL_WORD_MAP.keys()) + r")\b")
    text = pattern.sub(_replace_single, text)

    text = re.sub(r"\b([NSEW])\.(?=\s)", r"\1", text)
    text = re.sub(r"\b([NSEW]{1,2})[\.,;:]+(?=\s|$)", r"\1", text)
    return text


def _normalize_angles(text: str) -> str:
    text = re.sub(r"(?i)\bDEG(?:REE|REES)?\b", "°", text)
    text = re.sub(r"(?i)\bMIN(?:UTE|UTES)?\b", "'", text)
    text = re.sub(r"(?i)\bSEC(?:OND|ONDS)?\b", '"', text)
    text = re.sub(r"(?<=\d)\s*(?:°|º|o)\s*(?=\d)", "°", text, flags=re.IGNORECASE)
    text = re.sub(r"(?<=\d)\s*(?:'')\s*(?=\d)", '"', text)
    text = re.sub(r"(?<=\d)°(?=\d)", "° ", text)
    text = re.sub(r"(?<=\d)'(?=\d)", "' ", text)
    text = re.sub(r'(?<=\d)"(?=\d)', '" ', text)
    return text


def _uppercase_cue_words(text: str) -> str:
    for word in _CUE_WORDS:
        pattern = re.compile(rf"(?i)\b{word}\b")
        text = pattern.sub(word, text)
    return text


def clean_text(raw: str) -> str:
    """Return a normalized version of deed text for consistent parsing."""

    if not raw:
        return ""

    text = _normalize_characters(str(raw))
    text = _remove_headers_and_footers(text)
    text = _normalize_angles(text)
    text = _standardize_cardinals(text)
    text = _uppercase_cue_words(text)

    text = re.sub(r"\bTHEN\b", "THENCE", text, flags=re.IGNORECASE)
    text = re.sub(r"\b([NSEW]{1,2})\s*\.(?=\s|$)", r"\1", text)
    text = re.sub(r"(?<=\b\d)\s*[\.,](?=\s)", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


__all__ = ["clean_text"]

