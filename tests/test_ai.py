"""AI executive summary: facts, the figure check, retries, settings. The AI service is mocked."""
import json
import os
import stat

import pytest

from fpa import ai, pipeline


@pytest.fixture(scope="module")
def summary(tmp_path_factory):
    res = pipeline.run("2026-09", recalc=False, out_root=tmp_path_factory.mktemp("ai"))
    return res["summary"]


GOOD = ("Year to date, EBITDA was £114.6k against a budget of £134.3k, £19.7k below plan, with a margin of 18.3% "
        "against 20.7%. Let-only fees were £19.0k adverse on 81 lets. The full-year forecast is £194.8k against "
        "£239.3k. Cash stays above the £50.0k buffer, at £54.6k after corporation tax.")


def test_facts_contain_the_pack_figures(summary):
    f = ai.facts(summary)
    assert f["ytd"]["ebitda"]["actual"] == 114555 and f["ytd"]["ebitda"]["variance"] == -19748
    assert f["full_year_ebitda"]["budget"] == 239320 and f["cash"]["minimum_buffer"] == 50000
    assert any("Let-only fees" in c for c in f["ytd"]["commentary"])


def test_verify_accepts_true_figures_and_rejects_invented_ones(summary):
    f = ai.facts(summary)
    assert ai.verify(GOOD, f)
    with pytest.raises(ai.AIError, match="£25.0k"):
        ai.verify(GOOD.replace("£19.7k below", "£25.0k below"), f)
    with pytest.raises(ai.AIError, match="31.5%"):
        ai.verify(GOOD + " Margins improved to 31.5%.", f)


def test_draft_retries_once_then_succeeds(summary, monkeypatch):
    answers = iter([GOOD.replace("£194.8k", "£210.0k"), GOOD])
    prompts = []
    monkeypatch.setattr(ai, "call", lambda cfg, p: (prompts.append(p), next(answers))[1])
    res = ai.draft(summary, {"provider": "openai", "api_key": "x", "model": "m"})
    assert res["text"] == GOOD and res["checked"] and len(prompts) == 2
    assert "rejected" in prompts[1] and "£210.0k" in prompts[1]


def test_draft_gives_up_after_two_bad_drafts(summary, monkeypatch):
    monkeypatch.setattr(ai, "call", lambda cfg, p: "EBITDA was £999.9k.")
    with pytest.raises(ai.AIError, match="not in the pack"):
        ai.draft(summary, {"provider": "openai", "api_key": "x", "model": "m"})


def test_settings_are_local_and_private(tmp_path, monkeypatch):
    monkeypatch.setattr(ai, "SETTINGS", tmp_path / "home" / "ai.toml")
    monkeypatch.setattr(ai, "LEGACY", tmp_path / "local.toml")
    for v in ("OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
        monkeypatch.delenv(v, raising=False)
    assert ai.settings() is None
    with pytest.raises(ai.AIError, match="API key"):
        ai.save_settings("openai", "not a key")
    ai.save_settings("openai", "sk-test_" + "a" * 30, "gpt-test")
    assert ai.settings() == {"provider": "openai", "api_key": "sk-test_" + "a" * 30, "model": "gpt-test"}
    assert stat.S_IMODE(os.stat(tmp_path / "home" / "ai.toml").st_mode) == 0o600
    ai.clear_settings()
    assert ai.settings() is None
    monkeypatch.setenv("ANTHROPIC_API_KEY", "env-key")
    assert ai.settings()["provider"] == "anthropic" and ai.settings()["model"] == "claude-sonnet-5"


def test_request_shape_for_openai(monkeypatch):
    seen = {}

    def fake_post(url, headers, body, timeout=60):
        seen.update(url=url, headers=headers, body=body)
        return {"choices": [{"message": {"content": " Summary. "}}]}
    monkeypatch.setattr(ai, "_post", fake_post)
    assert ai.call({"provider": "openai", "api_key": "k", "model": "m"}, "hello") == "Summary."
    assert seen["url"].endswith("/v1/chat/completions") and seen["headers"]["Authorization"] == "Bearer k"
    assert json.dumps(seen["body"]).count("hello") == 1
