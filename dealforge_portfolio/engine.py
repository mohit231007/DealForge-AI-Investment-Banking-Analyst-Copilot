"""Sanitized DealForge portfolio engine.

The public edition intentionally uses structured, synthetic CSV inputs. It demonstrates
valuation-set QA, adjusted valuation triangulation, Office export automation, and
portable deal-pack generation without exposing the private Algosphere product core.
"""

from __future__ import annotations

import csv
import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches, Pt


@dataclass(frozen=True)
class PortfolioPackResult:
    output_dir: Path
    bundle_path: Path
    summary: dict[str, Any]
    artifacts: list[Path]


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def _number(value: Any) -> float | None:
    text = str(value or "").strip().replace(",", "")
    if not text or text.upper() in {"NM", "N/A", "NA", "NONE", "NULL", "-"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _clean_number(value: float | None) -> float | None:
    if value is None:
        return None
    return round(float(value), 2)


def _median(values: list[float | None]) -> float | None:
    clean = [float(value) for value in values if value is not None and value > 0]
    return _clean_number(median(clean)) if clean else None


def _normalize(value: Any) -> str:
    return " ".join(str(value or "").lower().replace("-", " ").replace("_", " ").split())


def _sector_fit(row_sector: Any, target_sector: Any) -> bool:
    row_tokens = set(_normalize(row_sector).split())
    target_tokens = set(_normalize(target_sector).split())
    return bool(row_tokens and target_tokens and row_tokens & target_tokens)


def _fmt(value: Any, unit: str = "") -> str:
    if value is None:
        return "N/A"
    if isinstance(value, float):
        text = f"{value:,.2f}".rstrip("0").rstrip(".")
    else:
        text = str(value)
    return f"{text} {unit}".strip()


def _metric_map(financial_rows: list[dict[str, str]]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for row in financial_rows:
        metric = str(row.get("metric") or "").strip()
        value = _number(row.get("value"))
        if not metric or value is None:
            continue
        result[metric.lower()] = {
            "metric": metric,
            "period": row.get("period") or "N/A",
            "value": value,
            "unit": row.get("unit") or "",
            "source": row.get("source") or "Synthetic portfolio input",
        }
    return result


def _audit_peers(company: str, ticker: str, market: str, sector: str, rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    reviewed: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        company_name = row.get("company") or f"Peer {index}"
        row_ticker = row.get("ticker") or ""
        ev_revenue = _number(row.get("ev_revenue_multiple"))
        ev_ebitda = _number(row.get("ev_ebitda_multiple"))
        is_target = _normalize(company_name) == _normalize(company) or (
            bool(ticker) and _normalize(row_ticker).split(".")[0] == _normalize(ticker).split(".")[0]
        )
        same_market = str(row.get("market") or "").upper() == market.upper()
        sector_fit = _sector_fit(row.get("sector"), sector)
        revenue_ok = not is_target and same_market and sector_fit and ev_revenue is not None and 0 < ev_revenue <= 25
        ebitda_ok = revenue_ok and ev_ebitda is not None and 0 < ev_ebitda <= 100
        reasons: list[str] = []
        if is_target:
            reasons.append("target company removed from primary peer median")
        if not same_market:
            reasons.append("market mismatch")
        if not sector_fit:
            reasons.append("sector mismatch")
        if ev_revenue is None or not 0 < ev_revenue <= 25:
            reasons.append("EV/Revenue missing or outside 0-25x review range")
        if revenue_ok and not ebitda_ok:
            reasons.append("EV/EBITDA is NM or outside 0-100x review range")
        reviewed.append(
            {
                **row,
                "row_index": index,
                "ev_revenue_multiple": ev_revenue,
                "ev_ebitda_multiple": ev_ebitda,
                "include_ev_revenue": revenue_ok,
                "include_ev_ebitda": ebitda_ok,
                "review_status": "INCLUDED" if revenue_ok else "EXCLUDED",
                "review_notes": "; ".join(reasons) if reasons else "included in adjusted peer set",
            }
        )
    return reviewed


def _audit_precedents(market: str, sector: str, rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    reviewed: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        ev_revenue = _number(row.get("ev_revenue_multiple"))
        ev_ebitda = _number(row.get("ev_ebitda_multiple"))
        same_market = str(row.get("market") or "").upper() == market.upper()
        sector_fit = _sector_fit(row.get("sector"), sector)
        revenue_ok = same_market and sector_fit and ev_revenue is not None and 0 < ev_revenue <= 25
        ebitda_ok = revenue_ok and ev_ebitda is not None and 0 < ev_ebitda <= 100
        reasons: list[str] = []
        if not same_market:
            reasons.append("market mismatch")
        if not sector_fit:
            reasons.append("sector mismatch")
        if ev_revenue is None or not 0 < ev_revenue <= 25:
            reasons.append("EV/Revenue missing or outside 0-25x review range")
        if revenue_ok and not ebitda_ok:
            reasons.append("EV/EBITDA is NM or outside 0-100x review range")
        reviewed.append(
            {
                **row,
                "row_index": index,
                "transaction_value": _number(row.get("transaction_value")),
                "ev_revenue_multiple": ev_revenue,
                "ev_ebitda_multiple": ev_ebitda,
                "include_ev_revenue": revenue_ok,
                "include_ev_ebitda": ebitda_ok,
                "review_status": "INCLUDED" if revenue_ok else "EXCLUDED",
                "review_notes": "; ".join(reasons) if reasons else "included in adjusted precedent set",
            }
        )
    return reviewed


def _build_summary(
    company_row: dict[str, str],
    metrics: dict[str, dict[str, Any]],
    peers: list[dict[str, Any]],
    precedents: list[dict[str, Any]],
) -> dict[str, Any]:
    revenue = metrics.get("revenue", {}).get("value")
    ebitda = metrics.get("ebitda", {}).get("value")
    unit = metrics.get("revenue", {}).get("unit") or "currency units"

    raw_peer_revenue = _median([row.get("ev_revenue_multiple") for row in peers])
    raw_peer_ebitda = _median([row.get("ev_ebitda_multiple") for row in peers])
    adjusted_peer_revenue = _median([row.get("ev_revenue_multiple") for row in peers if row.get("include_ev_revenue")])
    adjusted_peer_ebitda = _median([row.get("ev_ebitda_multiple") for row in peers if row.get("include_ev_ebitda")])
    raw_precedent_revenue = _median([row.get("ev_revenue_multiple") for row in precedents])
    raw_precedent_ebitda = _median([row.get("ev_ebitda_multiple") for row in precedents])
    adjusted_precedent_revenue = _median(
        [row.get("ev_revenue_multiple") for row in precedents if row.get("include_ev_revenue")]
    )
    adjusted_precedent_ebitda = _median(
        [row.get("ev_ebitda_multiple") for row in precedents if row.get("include_ev_ebitda")]
    )

    implied = {
        "comps_revenue_method": _clean_number(revenue * adjusted_peer_revenue)
        if revenue and adjusted_peer_revenue
        else None,
        "comps_ebitda_method": _clean_number(ebitda * adjusted_peer_ebitda)
        if ebitda and adjusted_peer_ebitda
        else None,
        "precedent_revenue_method": _clean_number(revenue * adjusted_precedent_revenue)
        if revenue and adjusted_precedent_revenue
        else None,
        "precedent_ebitda_method": _clean_number(ebitda * adjusted_precedent_ebitda)
        if ebitda and adjusted_precedent_ebitda
        else None,
    }
    clean_implied = [value for value in implied.values() if value is not None]
    transaction_units = {
        str(row.get("transaction_value_unit") or "").strip().lower()
        for row in precedents
        if row.get("include_ev_revenue")
    }
    source_review = any("synthetic" in str(row.get("source") or "").lower() for row in [*peers, *precedents])

    return {
        "status": "PASS_WITH_REVIEW_ITEMS",
        "company": company_row.get("company"),
        "ticker": company_row.get("ticker"),
        "market": company_row.get("market"),
        "sector": company_row.get("sector"),
        "target_revenue": revenue,
        "target_ebitda": ebitda,
        "target_unit": unit,
        "original_medians": {
            "comps_ev_revenue": raw_peer_revenue,
            "comps_ev_ebitda": raw_peer_ebitda,
            "precedent_ev_revenue": raw_precedent_revenue,
            "precedent_ev_ebitda": raw_precedent_ebitda,
        },
        "adjusted_medians": {
            "comps_ev_revenue": adjusted_peer_revenue,
            "comps_ev_ebitda": adjusted_peer_ebitda,
            "precedent_ev_revenue": adjusted_precedent_revenue,
            "precedent_ev_ebitda": adjusted_precedent_ebitda,
        },
        "adjusted_implied_enterprise_values": {
            **implied,
            "range": {
                "low": min(clean_implied) if clean_implied else None,
                "high": max(clean_implied) if clean_implied else None,
            },
            "unit": unit,
        },
        "included_peer_count": sum(bool(row.get("include_ev_revenue")) for row in peers),
        "excluded_peer_count": sum(not bool(row.get("include_ev_revenue")) for row in peers),
        "included_precedent_count": sum(bool(row.get("include_ev_revenue")) for row in precedents),
        "excluded_precedent_count": sum(not bool(row.get("include_ev_revenue")) for row in precedents),
        "source_verification_status": "NEEDS_REVIEW" if source_review else "PASS",
        "unit_normalization_status": "NEEDS_NORMALIZATION" if len(transaction_units) > 1 else "PASS",
        "peer_review_rows": peers,
        "precedent_review_rows": precedents,
        "disclaimer": "Portfolio demonstration only. Human review required. Not investment advice.",
    }


def _write_markdown(path: Path, content: str) -> Path:
    path.write_text(content.strip() + "\n", encoding="utf-8")
    return path


def _write_json(path: Path, payload: dict[str, Any]) -> Path:
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def _write_company_profile(path: Path, summary: dict[str, Any], metrics: dict[str, dict[str, Any]]) -> Path:
    metric_lines = "\n".join(
        f"- **{item['metric']}:** {_fmt(item['value'], item['unit'])} ({item['period']}; {item['source']})"
        for item in metrics.values()
    )
    content = f"""
# Company Profile

**Company:** {summary['company']}  
**Ticker:** {summary['ticker']}  
**Market:** {summary['market']}  
**Sector:** {summary['sector']}

## Selected financial metrics

{metric_lines}

## Valuation QA posture

- Included peer rows: {summary['included_peer_count']}
- Excluded peer rows: {summary['excluded_peer_count']}
- Included precedent rows: {summary['included_precedent_count']}
- Excluded precedent rows: {summary['excluded_precedent_count']}
- Source verification: {summary['source_verification_status']}
- Unit normalization: {summary['unit_normalization_status']}

This is a synthetic portfolio company. Human review required. Not investment advice.
"""
    return _write_markdown(path, content)


def _write_memo(path: Path, summary: dict[str, Any]) -> Path:
    adjusted = summary["adjusted_medians"]
    implied = summary["adjusted_implied_enterprise_values"]
    ev_range = implied["range"]
    content = f"""
# Investment Banking Memo

## Executive conclusion

DealForge AI produced a review-ready valuation scaffold for **{summary['company']}**. The adjusted
valuation set retains {summary['included_peer_count']} peers and {summary['included_precedent_count']}
precedent transactions after removing target/self, off-sector, and invalid multiple rows.

## Adjusted valuation summary

| Method | Adjusted multiple / output |
|---|---:|
| Comps EV/Revenue | {_fmt(adjusted['comps_ev_revenue'], 'x')} |
| Comps EV/EBITDA | {_fmt(adjusted['comps_ev_ebitda'], 'x')} |
| Precedent EV/Revenue | {_fmt(adjusted['precedent_ev_revenue'], 'x')} |
| Precedent EV/EBITDA | {_fmt(adjusted['precedent_ev_ebitda'], 'x')} |
| Adjusted EV range | {_fmt(ev_range['low'])} – {_fmt(ev_range['high'])} {implied['unit']} |

## Review posture

- Source verification: **{summary['source_verification_status']}**
- Transaction-unit normalization: **{summary['unit_normalization_status']}**
- The difference between original and adjusted medians is an explicit QA output, not an invisible override.

## Recommendation

Proceed only to focused diligence. Replace synthetic sources with verified filings, reconcile enterprise-value
bridges, normalize transaction values, and obtain analyst approval before external use.

**Portfolio demonstration only. Human review required. Not investment advice.**
"""
    return _write_markdown(path, content)


def _write_diligence(path: Path, summary: dict[str, Any]) -> Path:
    peer_exclusions = [row for row in summary["peer_review_rows"] if not row["include_ev_revenue"]]
    precedent_exclusions = [row for row in summary["precedent_review_rows"] if not row["include_ev_revenue"]]
    peer_lines = "\n".join(f"- {row.get('company')}: {row['review_notes']}" for row in peer_exclusions) or "- None"
    precedent_lines = (
        "\n".join(f"- {row.get('target')} / {row.get('acquirer')}: {row['review_notes']}" for row in precedent_exclusions)
        or "- None"
    )
    content = f"""
# Due Diligence Request List

## Financial validation

- Reconcile Revenue of {_fmt(summary['target_revenue'], summary['target_unit'])} to an audited statement.
- Reconcile EBITDA of {_fmt(summary['target_ebitda'], summary['target_unit'])} and document all adjustments.
- Confirm consolidated versus standalone reporting basis.
- Confirm net debt/cash, share count, and valuation date.

## Peer exclusions to review

{peer_lines}

## Precedent exclusions to review

{precedent_lines}

## Required next steps

- Replace synthetic source notes with verified URLs or document references.
- Normalize transaction values into one currency and scale.
- Validate business-model, geography, growth, and margin comparability.
- Review Excel formulas and assumptions before committee circulation.
- Obtain human analyst approval for all generated commentary.

**Portfolio demonstration only. Human review required. Not investment advice.**
"""
    return _write_markdown(path, content)


def _style_workbook(workbook: Workbook) -> None:
    navy = PatternFill("solid", fgColor="0B1F33")
    pale = PatternFill("solid", fgColor="EAF3F8")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(*(Side(style="thin", color="D9E2EC") for _ in range(4)))
    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.fill = navy
            cell.font = white_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if cell.row > 1 and cell.column == 1:
                    cell.fill = pale
        for column_cells in sheet.columns:
            letter = get_column_letter(column_cells[0].column)
            longest = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[letter].width = min(max(longest + 2, 12), 48)


def _write_excel(path: Path, summary: dict[str, Any], metrics: dict[str, dict[str, Any]]) -> Path:
    workbook = Workbook()
    cover = workbook.active
    cover.title = "01_Cover"
    cover.append(["DealForge AI Portfolio Valuation Model", "Value"])
    for label, value in [
        ("Company", summary["company"]),
        ("Ticker", summary["ticker"]),
        ("Market", summary["market"]),
        ("Sector", summary["sector"]),
        ("Status", summary["status"]),
        ("Source verification", summary["source_verification_status"]),
        ("Unit normalization", summary["unit_normalization_status"]),
    ]:
        cover.append([label, value])

    valuation = workbook.create_sheet("02_Valuation_Summary")
    valuation.append(["Metric", "Original", "Adjusted", "Unit"])
    original = summary["original_medians"]
    adjusted = summary["adjusted_medians"]
    for label, key in [
        ("Comps EV/Revenue", "comps_ev_revenue"),
        ("Comps EV/EBITDA", "comps_ev_ebitda"),
        ("Precedent EV/Revenue", "precedent_ev_revenue"),
        ("Precedent EV/EBITDA", "precedent_ev_ebitda"),
    ]:
        valuation.append([label, original[key], adjusted[key], "x"])
    implied = summary["adjusted_implied_enterprise_values"]
    valuation.append(["Adjusted EV range low", None, implied["range"]["low"], implied["unit"]])
    valuation.append(["Adjusted EV range high", None, implied["range"]["high"], implied["unit"]])

    financial = workbook.create_sheet("03_Financials")
    financial.append(["Metric", "Period", "Value", "Unit", "Source"])
    for item in metrics.values():
        financial.append([item["metric"], item["period"], item["value"], item["unit"], item["source"]])

    peer_sheet = workbook.create_sheet("04_Peer_QA")
    peer_sheet.append(["Company", "Ticker", "EV/Revenue", "EV/EBITDA", "Revenue set", "EBITDA set", "Review notes"])
    for row in summary["peer_review_rows"]:
        peer_sheet.append(
            [row.get("company"), row.get("ticker"), row.get("ev_revenue_multiple"), row.get("ev_ebitda_multiple"), row["include_ev_revenue"], row["include_ev_ebitda"], row["review_notes"]]
        )

    precedent_sheet = workbook.create_sheet("05_Precedent_QA")
    precedent_sheet.append(["Target", "Acquirer", "Year", "EV/Revenue", "EV/EBITDA", "Revenue set", "Review notes"])
    for row in summary["precedent_review_rows"]:
        precedent_sheet.append(
            [row.get("target"), row.get("acquirer"), row.get("announced_year"), row.get("ev_revenue_multiple"), row.get("ev_ebitda_multiple"), row["include_ev_revenue"], row["review_notes"]]
        )

    _style_workbook(workbook)
    workbook.save(path)
    return path


def _add_slide(prs: Presentation, title: str, lines: list[str]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(28)
    title_frame.paragraphs[0].font.bold = True
    body = slide.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.6), Inches(5.6))
    frame = body.text_frame
    frame.clear()
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.font.size = Pt(18)
        paragraph.space_after = Pt(10)


def _write_ppt(path: Path, summary: dict[str, Any]) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    _add_slide(prs, "DealForge AI — Portfolio Edition", [summary["company"], f"{summary['market']} | {summary['sector']}", "Built by Mohit Bhatnagar", "Human review required. Not investment advice."])
    _add_slide(prs, "Workflow", ["Structured financial inputs", "Peer and precedent QA", "Adjusted valuation triangulation", "Memo, Excel, PowerPoint, diligence list, and ZIP bundle"])
    _add_slide(prs, "Valuation QA", [f"Peers retained: {summary['included_peer_count']}", f"Precedents retained: {summary['included_precedent_count']}", f"Source status: {summary['source_verification_status']}", f"Unit status: {summary['unit_normalization_status']}"])
    adjusted = summary["adjusted_medians"]
    _add_slide(prs, "Adjusted Multiples", [f"Comps EV/Revenue: {_fmt(adjusted['comps_ev_revenue'], 'x')}", f"Comps EV/EBITDA: {_fmt(adjusted['comps_ev_ebitda'], 'x')}", f"Precedent EV/Revenue: {_fmt(adjusted['precedent_ev_revenue'], 'x')}"])
    ev_range = summary["adjusted_implied_enterprise_values"]["range"]
    _add_slide(prs, "Adjusted EV Range", [f"Low: {_fmt(ev_range['low'], summary['target_unit'])}", f"High: {_fmt(ev_range['high'], summary['target_unit'])}", "Directional review workpaper only"])
    _add_slide(prs, "Why this project matters", ["Finance-domain workflow design", "Python product engineering", "Office automation", "Red-team QA and visible review controls", "Business and technical communication"])
    prs.save(path)
    return path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def generate_portfolio_pack(input_dir: Path | str = "sample_data", output_dir: Path | str = "outputs/portfolio_demo") -> PortfolioPackResult:
    """Generate the sanitized portfolio deal pack from three synthetic CSV inputs."""

    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    company_rows = _read_csv(input_path / "company_financials.csv")
    peer_rows = _read_csv(input_path / "peer_comps.csv")
    precedent_rows = _read_csv(input_path / "precedent_transactions.csv")
    if not company_rows:
        raise ValueError("company_financials.csv must contain at least one row")

    company_row = company_rows[0]
    metrics = _metric_map(company_rows)
    company = company_row.get("company") or "Synthetic Company"
    ticker = company_row.get("ticker") or "SYNTH"
    market = company_row.get("market") or "IN"
    sector = company_row.get("sector") or "Consumer Internet"
    peers = _audit_peers(company, ticker, market, sector, peer_rows)
    precedents = _audit_precedents(market, sector, precedent_rows)
    summary = _build_summary(company_row, metrics, peers, precedents)

    artifacts = [
        _write_company_profile(output_path / "01_company_profile.md", summary, metrics),
        _write_memo(output_path / "02_investment_memo.md", summary),
        _write_json(output_path / "03_adjusted_valuation_summary.json", summary),
        _write_diligence(output_path / "04_due_diligence_request_list.md", summary),
        _write_excel(output_path / "05_valuation_model.xlsx", summary, metrics),
        _write_ppt(output_path / "06_pitchbook.pptx", summary),
    ]
    manifest = {
        "portfolio_edition": True,
        "company": company,
        "files": [
            {"name": artifact.name, "size_bytes": artifact.stat().st_size, "sha256": _sha256(artifact)}
            for artifact in artifacts
        ],
        "disclaimer": summary["disclaimer"],
    }
    manifest_path = _write_json(output_path / "07_manifest.json", manifest)
    artifacts.append(manifest_path)

    bundle_path = output_path / "dealforge_portfolio_pack.zip"
    with ZipFile(bundle_path, "w", ZIP_DEFLATED) as bundle:
        for artifact in artifacts:
            bundle.write(artifact, arcname=artifact.name)

    return PortfolioPackResult(output_dir=output_path, bundle_path=bundle_path, summary=summary, artifacts=artifacts)
