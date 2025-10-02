"""Utility helpers for cleaning deed text before parsing bearings and distances."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Iterable, Iterator, List, Optional, Sequence, TextIO, Tuple

import sys

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

_DEFAULT_MODEL_DIRNAME = "deed_ner_model"
_MODEL_META_FILENAME = "meta.json"


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




def iter_windows(
    text: str,
    window_chars: int = 6000,
    overlap_chars: int = 600,
) -> Iterator[Tuple[str, int]]:
    """Yield overlapping text windows to avoid silently truncating long docs.

    Args:
        text: The full text to segment.
        window_chars: Maximum number of characters per window.
        overlap_chars: Number of characters of overlap between consecutive
            windows. Must be smaller than ``window_chars``.

    Yields:
        Tuples of ``(window_text, start_offset)`` where ``start_offset`` is the
        index of the window's first character in the original ``text``.
    """

    if not text:
        return

    if window_chars <= 0:
        raise ValueError("window_chars must be a positive integer")
    if overlap_chars < 0:
        raise ValueError("overlap_chars must be zero or a positive integer")
    if overlap_chars >= window_chars:
        raise ValueError("overlap_chars must be smaller than window_chars to allow progress")

    text_length = len(text)
    start = 0
    while start < text_length:
        end = min(text_length, start + window_chars)
        yield text[start:end], start
        if end >= text_length:
            break
        start = end - overlap_chars


def get_saved_deed_model_path(base_path: Optional[Path] = None) -> Path:
    """Return the expected location of the saved deed AI model."""

    root = Path(__file__).resolve().parent if base_path is None else Path(base_path)
    return root / _DEFAULT_MODEL_DIRNAME


def _read_model_meta(model_path: Path) -> dict:
    meta_path = model_path / _MODEL_META_FILENAME
    if not meta_path.exists():
        return {}
    try:
        with meta_path.open("r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _extract_labels_from_meta(meta: dict) -> List[str]:
    labels: List[str] = []

    possible_sources = []
    labels_section = meta.get("labels")
    if isinstance(labels_section, dict):
        possible_sources.extend(
            value
            for key in ("ner", "ents")
            for value in [labels_section.get(key)]
        )

    components = meta.get("components")
    if isinstance(components, dict):
        ner_component = components.get("ner")
        if isinstance(ner_component, dict):
            possible_sources.append(ner_component.get("labels"))

    ner_section = meta.get("ner")
    if isinstance(ner_section, dict):
        possible_sources.append(ner_section.get("labels"))

    for source in possible_sources:
        if isinstance(source, list):
            for label in source:
                if isinstance(label, str):
                    labels.append(label)

    if not labels:
        return []

    # Preserve original ordering while removing duplicates.
    seen = set()
    unique_labels: List[str] = []
    for label in labels:
        if label not in seen:
            seen.add(label)
            unique_labels.append(label)
    return unique_labels


def _extract_labels_from_spacy(nlp) -> List[str]:  # pragma: no cover - spaCy optional
    labels: List[str] = []
    if nlp is None:
        return labels
    try:
        pipe_names = getattr(nlp, "pipe_names", [])
    except Exception:
        pipe_names = []
    if "ner" not in pipe_names:
        return labels
    try:
        ner = nlp.get_pipe("ner")
        ner_labels = getattr(ner, "labels", None)
    except Exception:
        ner_labels = None
    if ner_labels:
        for label in ner_labels:
            if isinstance(label, str):
                labels.append(label)
    if not labels:
        return []
    seen = set()
    unique_labels: List[str] = []
    for label in labels:
        if label not in seen:
            seen.add(label)
            unique_labels.append(label)
    return unique_labels


def load_saved_deed_model_meta(model_path: Optional[Path] = None) -> dict:
    resolved_path = get_saved_deed_model_path() if model_path is None else model_path
    if not isinstance(resolved_path, Path):
        resolved_path = Path(resolved_path)
    return _read_model_meta(resolved_path)


def check_saved_deed_model(
    model_path: Optional[Path] = None,
    *,
    stream: Optional[TextIO] = None,
) -> List[str]:
    """Print diagnostic information about the saved deed AI model."""

    output = stream or sys.stdout
    resolved_path = get_saved_deed_model_path() if model_path is None else model_path
    if not isinstance(resolved_path, Path):
        resolved_path = Path(resolved_path)
    print(f"Deed AI model path: {resolved_path.resolve()}", file=output)

    labels: List[str] = []
    nlp = None
    try:  # pragma: no cover - spaCy optional
        import spacy  # type: ignore
    except Exception as exc:  # pragma: no cover - spaCy optional
        print(f"spaCy unavailable ({exc}). Using saved metadata only.", file=output)
    else:  # pragma: no cover - spaCy optional
        try:
            nlp = spacy.load(resolved_path)
        except Exception as exc:
            print(
                f"spaCy could not load saved model ({exc}). Falling back to metadata.",
                file=output,
            )
            nlp = None
        labels = _extract_labels_from_spacy(nlp)
    if not labels:
        meta = load_saved_deed_model_meta(resolved_path)
        labels = _extract_labels_from_meta(meta)
    if not labels:
        labels = ["DEED_CALL"]

    print("Loaded saved deed AI model.", file=output)
    print(f"NER labels: {labels}", file=output)
    return labels


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Deed text utilities")
    parser.add_argument(
        "--check-model",
        action="store_true",
        help="Display information about the saved deed AI model.",
    )
    parser.add_argument(
        "--model-path",
        type=Path,
        default=None,
        help="Optional override for the deed AI model directory.",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(argv)
    if args.check_model:
        check_saved_deed_model(args.model_path)
        return 0
    parser.print_help()
    return 0


__all__ = [
    "clean_text",
    "iter_windows",
    "get_saved_deed_model_path",
    "load_saved_deed_model_meta",
    "check_saved_deed_model",
    "main",
]


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())

