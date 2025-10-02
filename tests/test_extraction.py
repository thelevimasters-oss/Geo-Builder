from __future__ import annotations

import pandas as pd
import pytest

import deed_extractor


class FakeSpan:
    def __init__(self, start: int, end: int, label: str = "DEED_CALL") -> None:
        self.start_char = start
        self.end_char = end
        self.label_ = label


class FakeDoc:
    def __init__(self, text: str, spans: list[FakeSpan]) -> None:
        self.text = text
        self.ents = spans


class FakeNLP:
    def __init__(self, trigger: str) -> None:
        self._trigger = trigger

    def __call__(self, text: str) -> FakeDoc:
        start = text.index(self._trigger)
        end = start + len(self._trigger)
        return FakeDoc(text, [FakeSpan(start, end)])


@pytest.fixture(autouse=True)
def reset_nlp_cache():
    """Ensure the cached spaCy pipeline is cleared between tests."""

    deed_extractor._NLP_CACHE = None
    yield
    deed_extractor._NLP_CACHE = None


def test_update_deed_model_cache_replaces_cache() -> None:
    sentinel = object()

    deed_extractor.update_deed_model_cache(sentinel)
    assert deed_extractor._NLP_CACHE is sentinel

    deed_extractor.update_deed_model_cache()
    assert deed_extractor._NLP_CACHE is None


def test_extract_calls_hybrid_uses_entity_ruler(monkeypatch: pytest.MonkeyPatch) -> None:
    raw_call = "Thence north 10 degrees east 120 feet."
    cleaned_call = deed_extractor.clean_text(raw_call)
    fake_nlp = FakeNLP(cleaned_call)

    monkeypatch.setattr(deed_extractor, "_get_deed_nlp", lambda: fake_nlp)
    monkeypatch.setattr(deed_extractor, "_should_use_ner", lambda: True)

    results = deed_extractor.extract_calls_hybrid(raw_call, window_chars=64, overlap_chars=16)

    assert [row["source"] for row in results] == ["ner"]
    assert results[0]["text"] == cleaned_call


def test_extract_calls_hybrid_regex_fallback(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(deed_extractor, "_should_use_ner", lambda: False)
    text = "Thence north 120 feet."

    results = deed_extractor.extract_calls_hybrid(text, window_chars=32, overlap_chars=8)

    assert results[0]["source"] == "regex"
    expected = deed_extractor.clean_text(text).rstrip(".")
    assert results[0]["text"].startswith(expected)


def test_extract_calls_hybrid_merges_overlapping_windows(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(deed_extractor, "_should_use_ner", lambda: False)

    def fake_iter_windows(text: str, *, window_chars: int, overlap_chars: int):
        yield text, 0
        yield text, 0

    def fake_iter_regex(window: str, *, offset: int):
        yield offset, offset + len(window)

    monkeypatch.setattr(deed_extractor, "iter_windows", fake_iter_windows)
    monkeypatch.setattr(deed_extractor, "_iter_regex_matches", fake_iter_regex)

    results = deed_extractor.extract_calls_hybrid("Thence north 120 feet.", window_chars=32, overlap_chars=8)

    assert len(results) == 1
    assert results[0]["source"] == "regex"


def test_canonicalize_df_start_end_labels() -> None:
    df = pd.DataFrame(
        {
            "DocID": ["doc", "doc"],
            "RawCall": ["Beginning at a point", "Thence west to the beginning"],
            "SourcePage": [1, 1],
            "CharStart": [0, 10],
            "CharEnd": [5, 25],
            "Extractor": ["regex", "regex"],
        }
    )

    canonical = deed_extractor.canonicalize_df(df)

    assert list(canonical["Start_End"]) == ["BEGIN", "END"]
    assert list(canonical["Sequence"]) == [1, 2]


def test_normalize_distance_handles_decimal_values() -> None:
    result = deed_extractor.normalize_distance("15.25 ft.")
    assert result == pytest.approx(15.25)


def test_normalize_distance_preserves_fractional_values() -> None:
    half_foot = deed_extractor.normalize_distance("1/2 ft")
    assert half_foot == pytest.approx(0.5)

    mixed_fraction = deed_extractor.normalize_distance("28 1/2 rods")
    assert mixed_fraction == pytest.approx(470.25)
