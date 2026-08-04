from evaluations.run_complete_transformer_paper_workflow import (
    has_incorrect_paper_location,
    has_incorrect_transformer_instruction,
    has_later_model_claim,
    has_unsupported_paper_claim,
    revised_quiz_ready,
    slides_artifact_ready,
)


SOURCE_ID = "SRC-123456789ABC"


def quiz_response(answer_key: str) -> dict:
    return {
        "artifact": {
            "has_file": True,
            "file_format": "docx",
            "content": (
                "# Transformer Quiz\n\n"
                "## Question 2\n\nCompare encoder and masked decoder attention to future tokens.\n\n"
                "## Answer Key\n\n" + answer_key
            ),
            "tool_trace": {"source_ids": [SOURCE_ID]},
        }
    }


def test_revised_transformer_quiz_requires_correct_causal_links():
    response = quiz_response(
        "Section 3.2.3 explains the autoregressive mask. T1 → T2 is prohibited, "
        "T1 → T3 is prohibited, and T2 → T3 is prohibited. T2 → T1 is allowed, and "
        "T3 → T2 is allowed. Equation 1 gives an attention result of [1.34, 0.99]. "
        "The worked calculation has q·k1 = 1, q·k2 = 0, and softmax weights 0.67 and 0.33. "
        "The learned and sinusoidal positional encodings performed nearly identical; the "
        "sinusoids may allow extrapolation."
    )

    assert revised_quiz_ready(response, SOURCE_ID)


def test_revised_transformer_quiz_rejects_future_link_as_allowed():
    response = quiz_response(
        "T1 → T2 is allowed. T2 → T3 is prohibited. T2 → T1 is allowed, and "
        "T3 → T2 is allowed. Equation 1 gives an attention result of [1.34, 0.99]. "
        "The worked calculation has q·k1 = 1, q·k2 = 0, and softmax weights 0.67 and 0.33. "
        "The learned and sinusoidal positional encodings performed nearly identical; the "
        "sinusoids may allow extrapolation."
    )

    assert not revised_quiz_ready(response, SOURCE_ID)


def test_transformer_workflow_rejects_wrong_paper_locations():
    assert has_incorrect_paper_location(
        "The sinusoidal positional encoding is Equation 1 in Section 3.5."
    )
    assert has_incorrect_paper_location(
        "The decoder mask is shown in Figure 2."
    )
    assert has_incorrect_paper_location(
        "Scaled dot-product attention is Equation 2 in the paper."
    )
    assert not has_incorrect_paper_location(
        "Equation 1 defines scaled dot-product attention; Section 3.5 gives positional encoding."
    )


def test_transformer_workflow_rejects_overstated_paper_claims():
    assert has_unsupported_paper_claim("The original Transformer was a proof-of-concept.")
    assert has_unsupported_paper_claim(
        "Learned positional embeddings cannot extrapolate to longer sequences."
    )
    assert not has_unsupported_paper_claim(
        "The authors chose sinusoids because they may allow extrapolation."
    )


def test_later_models_are_allowed_only_as_an_explicit_scope_boundary():
    assert not has_later_model_claim(
        "The paper does not discuss BERT, GPT, or other successors."
    )
    assert has_later_model_claim(
        "BERT later improved the Transformer by adding bidirectional training."
    )


def test_transformer_workflow_rejects_incorrect_mask_and_layer_instructions():
    assert has_incorrect_transformer_instruction(
        "Highlight the three sublayers per encoder block."
    )
    assert has_incorrect_transformer_instruction(
        "Multiply the mask matrix by the attention scores."
    )
    assert has_incorrect_transformer_instruction(
        "Use a lower-triangular matrix with zeros on and above the diagonal."
    )
    assert has_incorrect_transformer_instruction(
        "The ablation in Table 1 shows reduced path length."
    )
    assert has_incorrect_transformer_instruction(
        "Mask matrix M has ones below and zeros above. Add M as a negative infinity bias."
    )
    assert not has_incorrect_transformer_instruction(
        "The encoder has two sublayers. Add the causal mask before softmax; allowed "
        "links are on and below the diagonal. Do not multiply by the mask."
    )


def test_transformer_slides_require_pptx_source_trace_and_core_content():
    response = {
        "artifact": {
            "has_file": True,
            "file_format": "pptx",
            "content": (
                "Encoder and decoder attention lesson. Add an additive mask before softmax. "
                "Section 3.5 covers positional encoding. Evidence and interpretation are distinct."
            ),
            "tool_trace": {"source_ids": [SOURCE_ID]},
        }
    }

    assert slides_artifact_ready(response, SOURCE_ID)
    response["artifact"]["file_format"] = "docx"
    assert not slides_artifact_ready(response, SOURCE_ID)
