"""Regression tests for the scorecard calc engine.

The template workbook (`data/scorecard/FO_Market_Scorecard_Polo_OneSheet_v5.xlsx`)
had two data-quality bugs when it was first verified against Excel's own
cached values: cell E8 held text ("Macro Regime") instead of a numeric
CyclePsych weight, and row 27's Group label was "Flow " (trailing space),
excluding it from the Flow group score. Both have since been fixed in the
fixture (E8=0, A27="Flow") — see `test_detects_known_sheet_issues` below for
the values this produces, and `_group_score`'s docstring in scorecard.py for
why group matching is whitespace-tolerant regardless.
"""

from pathlib import Path

import pytest

from vnstock_agent import scorecard as sc

FIXTURE = Path(__file__).parent.parent / "data" / "scorecard" / "FO_Market_Scorecard_Polo_OneSheet_v5.xlsx"


@pytest.fixture(scope="module")
def result():
    wb, ws = sc.load_workbook_sheet(str(FIXTURE))
    base = sc.read_base_inputs(ws)
    return sc.compute(ws, base, sc.AutoData())


def test_group_scores(result):
    assert result.group_scores["Value"] == pytest.approx(67.81666666666666)
    assert result.group_scores["Technical"] == pytest.approx(46.71615906855472)
    assert result.group_scores["Flow"] == pytest.approx(54.34109792142938)
    assert result.group_scores["Macro"] == pytest.approx(78.89145106861635)
    assert result.group_scores["CyclePsych"] == pytest.approx(68.40942028985505)


def test_overall_score_and_stance(result):
    # CyclePsych weight (E8) is 0, so Overall Score = Value*.35 + Technical*.2 + Flow*.2 + Macro*.25
    assert result.overall_score == pytest.approx(63.67014749848424)
    assert result.stance == "Trung tính (giữ barbell)"


def test_regimes_and_cycle_phase(result):
    assert result.macro_regime == "Loose"
    assert result.sentiment_regime == "Neutral"
    assert result.cycle_phase == "Early Accumulation"


def test_suggested_equity_weight_and_leverage(result):
    assert result.suggested_equity_weight == pytest.approx(0.916850123978173)
    assert result.leverage_note == "zero"


def test_allocation_table(result):
    by_layer = {a["layer"]: a for a in result.allocation}
    assert by_layer["Core Compounders (C)"]["pct_equity"] == pytest.approx(0.6)
    assert by_layer["Growth Catalysts (G)"]["pct_equity"] == pytest.approx(0.25)
    assert by_layer["Lynch/Fisher Layer (L)"]["pct_equity"] == pytest.approx(0.15)
    assert by_layer["Core Compounders (C)"]["pct_nav"] == pytest.approx(0.5501100743869037)
    assert by_layer["Growth Catalysts (G)"]["pct_nav"] == pytest.approx(0.22921253099454325)
    assert by_layer["Lynch/Fisher Layer (L)"]["pct_nav"] == pytest.approx(0.13752751859672596)
    assert by_layer["Cash / Defensive"]["pct_nav"] == pytest.approx(0.083149876021827)


def test_narrative(result):
    assert result.narrative.startswith("Mid-cycle")


def test_detects_known_sheet_issues(result):
    """Both bugs are fixed in the fixture, so there should be no warnings —
    but E8=0 is surfaced as an informational note, not a silent no-op."""
    assert result.warnings == []
    assert any("E8" in n for n in result.notes)


def test_group_matching_is_whitespace_tolerant():
    """Regression guard for the original 'Flow ' bug: a stray space in a
    row's Group cell must not silently exclude it from its group score."""
    Row = sc.IndicatorRow
    rows = [
        Row(row=1, group="Flow ", factor="a", current=1, weight=1, good=0, bad=10,
            lower_is_better=True, frequency="Weekly", source="manual", score10=5, score100=50),
        Row(row=2, group="Flow", factor="b", current=1, weight=1, good=0, bad=10,
            lower_is_better=True, frequency="Weekly", source="manual", score10=7, score100=70),
    ]
    assert sc._group_score(rows, "Flow") == pytest.approx(60.0)


def test_derived_row_formulas_match_pe_z_score(result):
    row17 = next(r for r in result.rows if r.row == 17)
    # C17 = C14 - D14 = 14.9 - 16.2
    assert row17.current == pytest.approx(-1.3)
    row44 = next(r for r in result.rows if r.row == 44)
    # C44 = (C14-D14)/2.3
    assert row44.current == pytest.approx(-1.3 / 2.3)


def test_score10_clamped_to_0_10(result):
    for row in result.rows:
        assert 0.0 <= row.score10 <= 10.0


def test_liquidity_band(result):
    assert result.liquidity_band == "15-25k"


def test_manual_update_checklist_groups_by_frequency(result):
    checklist = sc.manual_update_checklist(result.rows)
    bucketed_rows = {r.row for v in checklist.values() for r in v}
    assert bucketed_rows <= {r.row for r in result.rows}
    all_factors = {r.factor for v in checklist.values() for r in v}
    # Auto-fetched/derived-from-auto indicators must not appear in the manual checklist.
    for auto_factor in ("RSI (VNIndex, 14d)", "Index / MA50 (ratio) - timing", "20D Avg Liquidity (VND bn)"):
        assert auto_factor not in all_factors
    # A known manual, weekly-frequency indicator must appear in the Weekly bucket.
    weekly_factors = {r.factor for r in checklist["Weekly"]}
    assert "CoE (%)" in weekly_factors
    quarterly_factors = {r.factor for r in checklist["Quarterly"]}
    assert "GDP (%)" in quarterly_factors
