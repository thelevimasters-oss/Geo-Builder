"""Utility helpers for cleaning deed text before parsing bearings and distances."""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import os
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

_ENTITY_RULER_PATTERNS = [
    {
        "label": "DEED_CALL",
        "pattern": [
            {"TEXT": {"REGEX": "(?i)^thence$"}, "OP": "?"},
            {"TEXT": {"REGEX": "(?i)^[ns](?:[ew])?$"}},
            {"TEXT": {"REGEX": r"^(?:\\d+(?:\\.\\d+)?|°|º|o|'|\"|-)$"}, "OP": "*"},
            {"TEXT": {"REGEX": "(?i)^[ns](?:[ew])?$"}, "OP": "?"},
            {"TEXT": {"REGEX": r"^\\d+(?:\\.\\d+)?$"}},
            {
                "LOWER": {
                    "IN": [
                        "feet",
                        "foot",
                        "ft",
                        "meter",
                        "meters",
                        "m",
                        "chain",
                        "chains",
                        "rod",
                        "rods",
                        "yard",
                        "yards",
                        "vara",
                        "varas",
                        "link",
                        "links",
                    ]
                },
                "OP": "?",
            },
        ],
    },
    {
        "label": "DEED_CALL",
        "pattern": [
            {"TEXT": {"REGEX": "(?i)^thence$"}, "OP": "?"},
            {
                "LOWER": {
                    "IN": [
                        "north",
                        "south",
                        "east",
                        "west",
                        "northeast",
                        "northwest",
                        "southeast",
                        "southwest",
                    ]
                }
            },
            {"TEXT": {"REGEX": r"^(?:along|with|of)$"}, "OP": "*"},
            {"TEXT": {"REGEX": r"^\\d+(?:\\.\\d+)?$"}},
            {
                "LOWER": {
                    "IN": [
                        "feet",
                        "foot",
                        "ft",
                        "meter",
                        "meters",
                        "m",
                        "chain",
                        "chains",
                        "rod",
                        "rods",
                        "yard",
                        "yards",
                        "vara",
                        "varas",
                        "link",
                        "links",
                    ]
                },
                "OP": "?",
            },
        ],
    },
]

_REGEX_FALLBACK_PATTERN = re.compile(
    r"""
    (?ix)
    \b
    (?:THENCE|THEN)?\s*
    (?:
        [NS](?:[EW])?
        (?:
            \s*\d{1,3}
            (?:\s*[°ºo]\s*\d{1,2}
                (?:\s*['′](?:\s*\d{1,2})?(?:\s*(?:"|″|''))? )?
            )?
        )?
        \s*[EW]?
        |
        (?:NORTH|SOUTH|EAST|WEST)(?:EAST|WEST)?
    )
    [^A-Z0-9]{0,10}
    \d+(?:\.\d+)?
    \s*
    (?:FEET|FOOT|FT|METERS?|M|CHAINS?|LINKS?|RODS?|YARDS?|VARAS?)
    """,
    re.VERBOSE,
)

_SPACY_AVAILABLE = importlib.util.find_spec("spacy") is not None
spacy = importlib.import_module("spacy") if _SPACY_AVAILABLE else None
_NLP_CACHE = None
_ENTITY_RULER_NAME = "deed_call_ruler"


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


def expand_span_to_line(text: str, start: int, end: int, *, context_chars: int = 80) -> str:
    """Expand a span to the surrounding call line.

    Args:
        text: Source text containing the span.
        start: Start index of the span.
        end: End index of the span.
        context_chars: Number of characters to expand on each side when the
            span is not bounded by newlines.

    Returns:
        A string containing the surrounding line when the text contains
        newline-delimited lines. If no newline boundaries exist, a substring
        expanded by ``context_chars`` characters on each side is returned.
    """

    if not text:
        return ""

    text_length = len(text)
    start = max(0, min(start, text_length))
    end = max(start, min(end, text_length))

    line_start = text.rfind("\n", 0, start)
    line_end = text.find("\n", end)

    if line_start != -1 or line_end != -1:
        if line_start == -1:
            line_start = 0
        else:
            line_start += 1
        if line_end == -1:
            line_end = text_length
        return text[line_start:line_end].strip()

    window_start = max(0, start - context_chars)
    window_end = min(text_length, end + context_chars)
    return text[window_start:window_end].strip()


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


def _ensure_entity_ruler(nlp) -> None:  # pragma: no cover - spaCy optional
    if nlp is None:
        return
    existing = []
    try:
        existing = list(getattr(nlp, "pipe_names", []))
    except Exception:
        existing = []
    if _ENTITY_RULER_NAME in existing:
        return
    try:
        before = "ner" if "ner" in existing else None
        ruler = nlp.add_pipe("entity_ruler", name=_ENTITY_RULER_NAME, before=before)
    except Exception:
        return
    try:
        ruler.add_patterns(_ENTITY_RULER_PATTERNS)
    except Exception:
        pass


def _ensure_ner_label(nlp) -> None:  # pragma: no cover - spaCy optional
    if nlp is None:
        return
    try:
        pipe_names = getattr(nlp, "pipe_names", [])
    except Exception:
        pipe_names = []
    if "ner" not in pipe_names:
        return
    try:
        ner = nlp.get_pipe("ner")
    except Exception:
        return
    add_label = getattr(ner, "add_label", None)
    if not callable(add_label):
        return
    try:
        add_label("DEED_CALL")
    except Exception:
        pass


def _get_deed_nlp():  # pragma: no cover - spaCy optional
    global _NLP_CACHE
    if _NLP_CACHE is not None:
        return _NLP_CACHE
    if spacy is None:
        _NLP_CACHE = None
        return _NLP_CACHE
    model_path = get_saved_deed_model_path()
    nlp = None
    if model_path.exists():
        try:
            nlp = spacy.load(model_path)
        except Exception:
            nlp = None
    if nlp is None:
        try:
            nlp = spacy.blank("en")
        except Exception:
            nlp = None
    if nlp is not None:
        _ensure_entity_ruler(nlp)
        _ensure_ner_label(nlp)
    _NLP_CACHE = nlp
    return _NLP_CACHE


def _should_use_ner() -> bool:
    disable_value = os.getenv("DEED_EXTRACTOR_DISABLE_NER")
    if disable_value is None:
        return True
    disable_value = disable_value.strip().lower()
    return disable_value not in {"1", "true", "yes", "on"}


def _iter_regex_matches(window: str, *, offset: int):
    for match in _REGEX_FALLBACK_PATTERN.finditer(window):
        start = offset + match.start()
        end = offset + match.end()
        yield start, end


def extract_calls_hybrid(
    text: str,
    *,
    window_chars: int = 6000,
    overlap_chars: int = 600,
) -> List[dict]:
    """Locate deed calls using spaCy NER with a regex fallback."""

    cleaned = clean_text(text)
    if not cleaned:
        return []

    use_ner = _should_use_ner()
    nlp = _get_deed_nlp() if use_ner else None

    results: List[dict] = []
    seen_spans = set()

    for window_text, start_offset in iter_windows(
        cleaned, window_chars=window_chars, overlap_chars=overlap_chars
    ):
        window_spans: List[Tuple[int, int]] = []
        if nlp is not None:
            try:
                doc = nlp(window_text)
            except Exception:
                doc = None
            if doc is not None:
                for ent in getattr(doc, "ents", []):
                    if ent.label_ != "DEED_CALL":
                        continue
                    span_start = start_offset + int(ent.start_char)
                    span_end = start_offset + int(ent.end_char)
                    if span_end <= span_start:
                        continue
                    window_spans.append((span_start, span_end))

        if not window_spans:
            window_spans.extend(_iter_regex_matches(window_text, offset=start_offset))

        for span_start, span_end in window_spans:
            if (span_start, span_end) in seen_spans:
                continue
            seen_spans.add((span_start, span_end))
            span_text = cleaned[span_start:span_end].strip()
            if not span_text:
                continue
            results.append(
                {
                    "text": span_text,
                    "start": span_start,
                    "end": span_end,
                    "label": "DEED_CALL",
                }
            )

    results.sort(key=lambda item: item.get("start", 0))
    return results


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
    "extract_calls_hybrid",
    "main",
]


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())

