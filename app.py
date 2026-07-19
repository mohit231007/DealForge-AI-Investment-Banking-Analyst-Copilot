from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from dealforge_portfolio import generate_portfolio_pack


st.set_page_config(page_title="DealForge AI Portfolio", page_icon="📊", layout="wide")

st.title("DealForge AI — Investment Banking Analyst Copilot")
st.caption("Portfolio-plus public edition by Mohit Bhatnagar · Synthetic data · Human review required")

st.markdown(
    """
This public version now mirrors the visible depth of a real DealForge analyst workbench while staying synthetic and safe:
source pack, citation map, financial extract, comps, precedents, source confidence, valuation audit, adjustment recommendations,
corrected workpapers, Excel model, PowerPoint pitchbook, memos, diligence list, consistency report, and ZIP bundle.
"""
)

with st.sidebar:
    st.header("Input mode")
    use_samples = st.checkbox("Use included synthetic sample data", value=True)
    company_file = st.file_uploader("company_financials.csv", type="csv", disabled=use_samples)
    peer_file = st.file_uploader("peer_comps.csv", type="csv", disabled=use_samples)
    precedent_file = st.file_uploader("precedent_transactions.csv", type="csv", disabled=use_samples)
    generate = st.button("Regenerate portfolio-plus deal pack", type="primary", use_container_width=True)
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

st.success("Portfolio-plus deal pack generated")

metrics = st.columns(8)
metrics[0].metric("Artifacts", summary["artifact_count"])
metrics[1].metric("Peers retained", summary["included_peer_count"])
metrics[2].metric("Peers excluded", summary["excluded_peer_count"])
metrics[3].metric("Precedents retained", summary["included_precedent_count"])
metrics[4].metric("Audit score", audit["score"])
metrics[5].metric("Audit warnings", audit["warning_count"])
metrics[6].metric("Source status", summary["source_verification_status"])
metrics[7].metric("Unit status", summary["unit_normalization_status"])

st.subheader("Adjusted valuation range")
st.markdown(f"### {ev_range['low']:,.2f} – {ev_range['high']:,.2f} {summary['target_unit']}")
st.caption("Workflow demonstration only. Human review required. Not investment advice.")

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
"""
    )

with source_tab:
    st.subheader("Source index")
    st.dataframe(pd.DataFrame(summary["source_index"]["documents"]), use_container_width=True)
    st.subheader("Citation map preview")
    st.dataframe(pd.DataFrame(summary["source_index"]["documents"]), use_container_width=True)

with financial_tab:
    st.subheader("Selected financial metrics")
    st.dataframe(pd.DataFrame(summary["financial_extract"]["metrics"]), use_container_width=True)

with comps_tab:
    st.subheader("Comparable-company QA")
    st.dataframe(pd.DataFrame(summary["peer_review_rows"]), use_container_width=True)

with precedents_tab:
    st.subheader("Precedent-transaction QA")
    st.dataframe(pd.DataFrame(summary["precedent_review_rows"]), use_container_width=True)

with confidence_tab:
    st.subheader("Source confidence ladder")
    st.dataframe(pd.DataFrame(source_confidence["rows"]), use_container_width=True)

with audit_tab:
    st.subheader("Valuation input audit")
    st.dataframe(pd.DataFrame(audit["checks"]), use_container_width=True)
    st.subheader("Adjustment recommendations")
    st.dataframe(pd.DataFrame(recommendations["recommendations"]), use_container_width=True)

with consistency_tab:
    st.subheader("Cross-artifact consistency")
    st.dataframe(pd.DataFrame(consistency["checks"]), use_container_width=True)

with preview_tab:
    st.subheader("Adjusted valuation summary JSON")
    st.json(summary["adjusted_valuation_summary"])

with downloads_tab:
    st.download_button(
        "Download complete portfolio-plus deal pack",
        data=result.bundle_path.read_bytes(),
        file_name=result.bundle_path.name,
        mime="application/zip",
        use_container_width=True,
    )
    st.subheader("Generated artifacts")
    for artifact in result.artifacts:
        st.write(f"`{artifact.name}` · {artifact.stat().st_size:,} bytes")

st.divider()
st.caption("Portfolio demonstration only. Human review required. Not investment advice.")
