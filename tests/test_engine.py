from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook
from pptx import Presentation

from dealforge_portfolio import generate_portfolio_pack


def test_portfolio_pack_generates_adjusted_valuation_outputs(tmp_path: Path) -> None:
    result = generate_portfolio_pack("sample_data", tmp_path / "pack")
    summary = result.summary

    assert summary["company"] == "AstraFood Technologies Ltd"
    assert summary["included_peer_count"] == 4
    assert summary["excluded_peer_count"] == 1
    assert summary["included_precedent_count"] == 4
    assert summary["excluded_precedent_count"] == 1
    assert summary["adjusted_medians"]["comps_ev_revenue"] == 6.4
    assert summary["adjusted_medians"]["comps_ev_ebitda"] == 60.0
    assert summary["adjusted_medians"]["precedent_ev_revenue"] == 4.65
    assert summary["adjusted_implied_enterprise_values"]["range"] == {
        "low": 45000.0,
        "high": 80000.0,
    }
    assert summary["source_verification_status"] == "NEEDS_REVIEW"
    assert summary["unit_normalization_status"] == "NEEDS_NORMALIZATION"

    names = {path.name for path in result.artifacts}
    assert names == {
        "01_company_profile.md",
        "02_investment_memo.md",
        "03_adjusted_valuation_summary.json",
        "04_due_diligence_request_list.md",
        "05_valuation_model.xlsx",
        "06_pitchbook.pptx",
        "07_manifest.json",
    }


def test_office_outputs_open_and_bundle_contains_all_artifacts(tmp_path: Path) -> None:
    result = generate_portfolio_pack("sample_data", tmp_path / "pack")

    workbook = load_workbook(result.output_dir / "05_valuation_model.xlsx", data_only=False)
    assert workbook.sheetnames == [
        "01_Cover",
        "02_Valuation_Summary",
        "03_Financials",
        "04_Peer_QA",
        "05_Precedent_QA",
    ]
    assert workbook["02_Valuation_Summary"]["C2"].value == 6.4

    presentation = Presentation(result.output_dir / "06_pitchbook.pptx")
    assert len(presentation.slides) == 6
    deck_text = "\n".join(
        shape.text for slide in presentation.slides for shape in slide.shapes if hasattr(shape, "text")
    )
    assert "DealForge AI" in deck_text
    assert "Human review required" in deck_text

    with ZipFile(result.bundle_path) as bundle:
        archived = set(bundle.namelist())
    assert archived == {path.name for path in result.artifacts}


def test_text_outputs_do_not_leak_python_null_literals(tmp_path: Path) -> None:
    result = generate_portfolio_pack("sample_data", tmp_path / "pack")
    for name in [
        "01_company_profile.md",
        "02_investment_memo.md",
        "04_due_diligence_request_list.md",
    ]:
        text = (result.output_dir / name).read_text(encoding="utf-8")
        assert "None" not in text
        assert "human review" in text.lower()
        assert "not investment advice" in text.lower()
