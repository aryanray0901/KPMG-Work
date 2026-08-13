"""Offline checks for failure diagnosis and safe wording suggestions."""

from __future__ import annotations

import hashlib
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from pptx import Presentation

# The production package installs openai from requirements.txt. The offline
# packaging test supplies a tiny import stub and keeps OPENAI_API_KEY empty, so
# no network path is exercised.
import types
class _OfflineOpenAI:
    def __init__(self, *args, **kwargs):
        raise RuntimeError("offline test")
sys.modules.setdefault("openai", types.SimpleNamespace(OpenAI=_OfflineOpenAI))

from pptx_editor import diagnose_failure_message, _chat_error_plan, plan_edit


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    sample = ROOT / "sample_files" / "kpmg_advisory_q3_original.pptx"
    assert sample.exists(), sample
    count = len(Presentation(sample).slides)
    assert count >= 2

    old_key = os.environ.get("OPENAI_API_KEY")
    os.environ["OPENAI_API_KEY"] = ""
    try:
        before = digest(sample)
        message = diagnose_failure_message(
            str(sample),
            f"move slide {count + 4} to slide 2",
            1,
            [],
            f"Slide {count + 4} does not exist.",
            "Resolving the source slide",
        )
        assert message.startswith("error cant do that\n"), message
        assert "What failed:" in message
        assert "Why:" in message
        assert "Try:" in message
        assert "PowerPoint was not changed." in message
        assert str(count) in message

        ambiguous = diagnose_failure_message(
            str(sample),
            "move the summary slide to the end",
            1,
            [],
            "The title reference is ambiguous and matched multiple slides.",
            "Resolving the slide title",
        )
        assert "slide 1" in ambiguous.lower() or "exact title" in ambiguous.lower()

        # The model-assisted diagnosis path must preserve the exact task and
        # return the structured corrective wording when the API is available.
        import pptx_editor as pe
        original_openai = pe.OpenAI
        class _FakeResponses:
            def parse(self, **kwargs):
                return types.SimpleNamespace(output_parsed=pe.FailureDiagnosis(
                    failed_step="Moving the duplicated slide",
                    reason="The pronoun did not resolve to a unique slide.",
                    suggested_wording="Duplicate slide 4, then move the duplicated slide to the final position.",
                    alternate_wording="First duplicate slide 4. Then move the newly duplicated slide to the end.",
                ))
        class _FakeClient:
            def __init__(self, *args, **kwargs):
                self.responses = _FakeResponses()
        pe.OpenAI = _FakeClient
        os.environ["OPENAI_API_KEY"] = "test-key"
        try:
            assisted = pe.diagnose_failure_message(
                str(sample),
                "copy slide 4 and put it at the back",
                1,
                [],
                "Reference could not be resolved.",
                "Resolving it",
            )
            assert "Moving the duplicated slide" in assisted
            assert "Duplicate slide 4" in assisted
        finally:
            pe.OpenAI = original_openai
            os.environ["OPENAI_API_KEY"] = ""

        failed_plan = _chat_error_plan(
            "No usable edit operation was generated.",
            pptx_path=str(sample),
            user_message="make it better",
            selected_slide=2,
            chat_history=[],
            failed_step="Turning the request into PowerPoint operations",
        )
        assert failed_plan["failed"] is True
        assert failed_plan["operations"] == []
        assert failed_plan["message"].startswith("error cant do that")
        assert "Try:" in failed_plan["message"]

        # No-key planning failure must produce guidance and leave the deck byte-identical.
        invalid = plan_edit(str(sample), f"delete slide {count + 50}", 1, chat_history=[])
        assert invalid["failed"] is True
        assert invalid["operations"] == []
        assert invalid["message"].startswith("error cant do that")
        assert "Try:" in invalid["message"]
        assert digest(sample) == before
    finally:
        if old_key is None:
            os.environ.pop("OPENAI_API_KEY", None)
        else:
            os.environ["OPENAI_API_KEY"] = old_key

    print("Failure diagnostics: PASS")


if __name__ == "__main__":
    main()
