"""Regression tests for the scorecard calc engine.

These check the pure-Python formula replication in `vnstock_agent.scorecard`
against the cached values embedded in the original workbook by whatever
spreadsheet engine last computed it — i.e. they confirm `compute()` produces
byte-for-byte the same Group Scores / Overall Score / Stance / Cycle Phase /
Allocation numbers Excel would, given the same inputs and no live data fetch.
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
    assert result.group_scores["Flow"] == pytest.approx(56.73555800784733)
    assert result.group_scores["Macro"] == pytest.approx(78.89145106861635)
    assert result.group_scores["CyclePsych"] == pytest.approx(68.40942028985505)


def test_overall_score_and_stance(result):
    assert result.overall_score == pytest.approx(64.14903951576783)
    assert result.stance == "Trung tính (giữ barbell)"


def test_regimes_and_cycle_phase(result):
    assert result.macro_regime == "Loose"
    assert result.sentiment_regime == "Neutral"
    assert result.cycle_phase == "Early Accumulation"


def test_suggested_equity_weight_and_leverage(result):
    assert result.suggested_equity_weight == pytest.approx(0.9237461690270566)
    assert result.leverage_note == "zero"


def test_allocation_table(result):
    by_layer = {a["layer"]: a for a in result.allocation}
    assert by_layer["Core Compounders (C)"]["pct_equity"] == pytest.approx(0.6)
    assert by_layer["Growth Catalysts (G)"]["pct_equity"] == pytest.approx(0.25)
    assert by_layer["Lynch/Fisher Layer (L)"]["pct_equity"] == pytest.approx(0.15)
    assert by_layer["Core Compounders (C)"]["pct_nav"] == pytest.approx(0.5542477014162339)
    assert by_layer["Growth Catalysts (G)"]["pct_nav"] == pytest.approx(0.23093654225676416)
    assert by_layer["Lynch/Fisher Layer (L)"]["pct_nav"] == pytest.approx(0.1385619253540585)
    assert by_layer["Cash / Defensive"]["pct_nav"] == pytest.approx(0.07625383097294336)


def test_narrative(result):
    assert result.narrative.startswith("Mid-cycle")


def test_detects_known_sheet_issues(result):
    """The original workbook has two latent bugs (non-numeric E8 weight, and
    a trailing space on row 27's group label) that this engine surfaces
    instead of silently working around."""
    joined = " ".join(result.warnings)
    assert "E8" in joined
    assert "Row 27" in joined


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
