from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from dealforge_portfolio import generate_portfolio_pack


st.set_page_config(
    page_title="DealForge AI Portfolio",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    .block-container {padding-top: 2rem; max-width: 1280px;}
    .deal-hero {
        padding: 1.25rem 1.35rem;
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 18px;
        background: linear-gradient(135deg, rgba(11,31,51,0.85), rgba(18,85,70,0.32));
        margin-bottom: 1rem;
    }
    .deal-eyebrow {
        color: #99F6E4;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.09em;
        margin-bottom: 0.25rem;
    }
    .deal-title {
        color: #FFFFFF;
        font-size: 2.6rem;
        line-height: 1.08;
        font-weight: 850;
        margin: 0;
    }
    .deal-subtitle {
        color: rgba(255,255,255,0.78);
        font-size: 1.05rem;
        max-width: 980px;
        margin-top: 0.75rem;
    }
    .card-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.85rem;
        margin: 1rem 0 1.1rem 0;
    }
    .deal-card {
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 16px;
        padding: 1rem 1rem 0.95rem 1rem;
        background: rgba(255,255,255,0.035);
        min-height: 112px;
    }
    .deal-card-label {
        color: rgba(255,255,255,0.65);
        font-size: 0.82rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.055em;
    }
    .deal-card-value {
        color: #FFFFFF;
        font-size: 1.8rem;
        font-weight: 850;
        line-height: 1.15;
        margin-top: 0.25rem;
        word-break: normal;
    }
    .deal-card-note {
        color: rgba(255,255,255,0.62);
        font-size: 0.82rem;
        margin-top: 0.3rem;
    }
    .status-pill {
        display: inline-block;
        padding: 0.35rem 0.65rem;
        border-radius: 999px;
        font-weight: 800;
        font-size: 0.83rem;
        background: rgba(251,191,36,0.18);
        color: #FDE68A;
        border: 1px solid rgba(251,191,36,0.35);
        margin-top: 0.35rem;
    }
    .success-pill {
        display: inline-block;
        padding: 0.35rem 0.65rem;
        border-radius: 999px;
        font-weight: 800;
        font-size: 0.83rem;
        background: rgba(34,197,94,0.18);
        color: #BBF7D0;
        border: 1px solid rgba(34,197,94,0.35);
        margin-top: 0.35rem;
    }
    .banker-note {
        border-left: 4px solid #22C55E;
        padding: 0.85rem 1rem;
        background: rgba(34,197,94,0.08);
        border-radius: 12px;
        margin: 0.5rem 0 1rem 0;
        color: rgba(255,255,255,0.86);
    }
    @media (max-width: 900px) {
        .card-grid {grid-template-columns: repeat(2, minmax(0, 1fr));}
        .deal-title {font-size: 2rem;}
    }
</style>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Input mode")
    use_samples = st.checkbox("Use included synthetic sample data", value=True)
    company_file = st.file_uploader("company_financials.csv", type="csv", disabled=use_samples)
    peer_file = st.file_uploader("peer_comps.csv", type="csv", disabled=use_samples)
    precedent_file = st.file_uploader("precedent_transactions.csv", type="csv", disabled=use_samples)
    generate = st.button("Regenerate portfolio-plus deal pack", type="primary", width="stretch")
    st.divider()
    st.caption("Public portfolio repo: synthetic demo only. Private Algosphere product remains separate.")


@st.cache_data(show_spinner=False)
def _generate_from_samples() -> tuple[object, str]:
    workspace = Path(tempfile.mkdtemp(prefix="dealforge_portfolio_plus_"))
    result = generate_portfolio_pack(input_dir=Path("sample_data"), output_dir=workspace / "deal_pack")
    return result, str(workspace)


def _generate_from_uploads() -> object:
    workspace = Path(tempfile.mkdtemp(prefix="dealforge_portfolio_plus_"))
    if not all([company_file, peer_file, precedent_file]):
        st.error("Upload all three CSV files or use the included sample data.")
        st.stop()
    input_dir = workspace / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)
    for uploaded, name in [
        (company_file, "company_financials.csv"),
        (peer_file, "peer_comps.csv"),
        (precedent_file, "precedent_transactions.csv"),
    ]:
        (input_dir / name).write_bytes(uploaded.getvalue())
    return generate_portfolio_pack(input_dir=input_dir, output_dir=workspace / "deal_pack")


if use_samples and not generate:
    result, _workspace = _generate_from_samples()
elif use_samples and generate:
    _generate_from_samples.clear()
    result, _workspace = _generate_from_samples()
elif generate:
    result = _generate_from_uploads()
else:
    st.info("Upload all three CSV files or select the included synthetic sample data.")
    st.stop()

summary = result.summary
implied = summary["adjusted_implied_enterprise_values"]
ev_range = implied["range"]
audit = summary["valuation_input_audit"]
recommendations = summary["valuation_adjustment_recommendations"]
source_confidence = summary["source_confidence"]
consistency = summary["consistency_checks"]
source_index = summary.get("source_index", {"documents": []})
citation_map = summary.get("citation_map")
if not isinstance(citation_map, dict):
    citation_map = {
        "citations": [
            {
                "citation_id": row.get("document_id", f"DOC{index:03d}"),
                "file_name": row.get("file_name", row.get("name", "synthetic source")),
                "source_type": row.get("source_type", "synthetic_portfolio_source"),
                "preview": row.get("description", "Synthetic source row used for public portfolio demonstration."),
            }
            for index, row in enumerate(source_index.get("documents", []), start=1)
            if isinstance(row, dict)
        ]
    }

st.markdown(
    f"""
<div class="deal-hero">
    <div class="deal-eyebrow">Portfolio-plus public edition · Synthetic banker workbench</div>
    <h1 class="deal-title">DealForge AI — Investment Banking Analyst Copilot</h1>
    <div class="deal-subtitle">
        Built by Mohit Bhatnagar. A public-safe, synthetic finance automation workbench that compiles source pack,
        citation map, financial extract, comps, precedents, source confidence, valuation audit, adjustment recommendations,
        corrected workpapers, Excel model, PowerPoint pitchbook, memos, diligence list, consistency report, and ZIP bundle.
    </div>
</div>
<div class="banker-note">
    <strong>Recruiter / banker takeaway:</strong> DealForge generated a {summary['artifact_count']}-artifact deal pack,
    removed the target-company peer row, excluded the off-sector precedent, kept NM EBITDA rows out of EBITDA medians,
    and produced an adjusted EV range of <strong>{ev_range['low']:,.0f}–{ev_range['high']:,.0f} {summary['target_unit']}</strong>.
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    f"""
<div class="card-grid">
    <div class="deal-card">
        <div class="deal-card-label">Deal-pack artifacts</div>
        <div class="deal-card-value">{summary['artifact_count']}</div>
        <div class="deal-card-note">Excel, PPT, memos, QA JSON, CSVs, ZIP</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Adjusted EV range</div>
        <div class="deal-card-value">{ev_range['low']:,.0f}–{ev_range['high']:,.0f}</div>
        <div class="deal-card-note">{summary['target_unit']} · synthetic sample case</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Valuation QA status</div>
        <div class="deal-card-value">Review-ready</div>
        <div class="status-pill">{summary['adjusted_valuation_summary']['status']}</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Review flags</div>
        <div class="deal-card-value">{audit['warning_count']}</div>
        <div class="deal-card-note">source verification + unit normalization</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Peer-set decision</div>
        <div class="deal-card-value">{summary['included_peer_count']} kept / {summary['excluded_peer_count']} removed</div>
        <div class="deal-card-note">target-company row excluded</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Precedent decision</div>
        <div class="deal-card-value">{summary['included_precedent_count']} kept / {summary['excluded_precedent_count']} removed</div>
        <div class="deal-card-note">off-sector transaction excluded</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Source verification</div>
        <div class="deal-card-value">Needs review</div>
        <div class="status-pill">{summary['source_verification_status']}</div>
    </div>
    <div class="deal-card">
        <div class="deal-card-label">Unit normalization</div>
        <div class="deal-card-value">Needs review</div>
        <div class="status-pill">{summary['unit_normalization_status']}</div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

left, right = st.columns(2)
with left:
    st.subheader("Original vs adjusted medians")
    original = summary["original_medians"]
    adjusted = summary["adjusted_medians"]
    st.table(
        [
            {"Metric": "Comps EV/Revenue", "Original": original["comps_ev_revenue"], "Adjusted": adjusted["comps_ev_revenue"]},
            {"Metric": "Comps EV/EBITDA", "Original": original["comps_ev_ebitda"], "Adjusted": adjusted["comps_ev_ebitda"]},
            {"Metric": "Precedent EV/Revenue", "Original": original["precedent_ev_revenue"], "Adjusted": adjusted["precedent_ev_revenue"]},
            {"Metric": "Precedent EV/EBITDA", "Original": original["precedent_ev_ebitda"], "Adjusted": adjusted["precedent_ev_ebitda"]},
        ]
    )
with right:
    st.subheader("Adjusted implied enterprise values")
    st.table(
        [
            {"Method": key.replace("_", " ").title(), "Value": value, "Unit": implied["unit"]}
            for key, value in implied.items()
            if key not in {"range", "unit"}
        ]
    )

(
    dashboard_tab,
    source_tab,
    financial_tab,
    comps_tab,
    precedents_tab,
    confidence_tab,
    audit_tab,
    consistency_tab,
    preview_tab,
    downloads_tab,
) = st.tabs(
    [
        "Executive dashboard",
        "Source coverage",
        "Financial extraction",
        "Comps analysis",
        "Precedent transactions",
        "Source confidence",
        "Valuation QA",
        "Consistency checks",
        "Artifact preview",
        "Downloads",
    ]
)

with dashboard_tab:
    st.markdown(
        f"""
### {summary['company']} public portfolio case

- Market: **{summary['market']}**
- Ticker: **{summary['ticker']}**
- Sector: **{summary['sector']}**
- Revenue: **{summary['target_revenue']:,.0f} {summary['target_unit']}**
- EBITDA: **{summary['target_ebitda']:,.0f} {summary['target_unit']}**
- Generated output pack: **{summary['artifact_count']} artifacts + ZIP bundle**
- Professional-use posture: **human review required; not investment advice**
"""
    )

with source_tab:
    st.subheader("Source index")
    st.dataframe(pd.DataFrame(source_index.get("documents", [])), width="stretch")
    st.subheader("Citation map preview")
    st.dataframe(pd.DataFrame(citation_map.get("citations", [])), width="stretch")

with financial_tab:
    st.subheader("Selected financial metrics")
    st.dataframe(pd.DataFrame(summary["financial_extract"]["metrics"]), width="stretch")

with comps_tab:
    st.subheader("Comparable-company QA")
    st.dataframe(pd.DataFrame(summary["peer_review_rows"]), width="stretch")

with precedents_tab:
    st.subheader("Precedent-transaction QA")
    st.dataframe(pd.DataFrame(summary["precedent_review_rows"]), width="stretch")

with confidence_tab:
    st.subheader("Source confidence ladder")
    st.dataframe(pd.DataFrame(source_confidence["rows"]), width="stretch")

with audit_tab:
    st.subheader("Valuation input audit")
    st.dataframe(pd.DataFrame(audit["checks"]), width="stretch")
    st.subheader("Adjustment recommendations")
    st.dataframe(pd.DataFrame(recommendations["recommendations"]), width="stretch")

with consistency_tab:
    st.subheader("Cross-artifact consistency")
    st.dataframe(pd.DataFrame(consistency["checks"]), width="stretch")

with preview_tab:
    st.subheader("Adjusted valuation summary JSON")
    st.json(summary["adjusted_valuation_summary"])

with downloads_tab:
    st.download_button(
        "Download complete portfolio-plus deal pack",
        data=result.bundle_path.read_bytes(),
        file_name=result.bundle_path.name,
        mime="application/zip",
        width="stretch",
    )
    st.subheader("Generated artifacts")
    for artifact in result.artifacts:
        st.write(f"`{artifact.name}` · {artifact.stat().st_size:,} bytes")

st.divider()
st.caption("Portfolio demonstration only. Human review required. Not investment advice.")
