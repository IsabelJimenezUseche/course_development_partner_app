from evaluations.run_artifact_evals import normalize_source_identifier


def test_source_identifier_normalizes_model_typography() -> None:
    assert normalize_source_identifier("【SRC‑AB4E9A10D062】") == "【SRC-AB4E9A10D062】"


def test_source_identifier_preserves_identifier_characters() -> None:
    assert normalize_source_identifier("SRC-0785E188F4FB") == "SRC-0785E188F4FB"
