from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from dealforge_portfolio import PortfolioPackResult, generate_portfolio_pack


st.set_page_config(page_title="DealForge AI Portfolio", page_icon="📊", layout="wide")


def _generate_pack(input_dir: Path) -> PortfolioPackResult:
    workspace = Path(tempfile.mkdtemp(prefix="dealforge_portfolio_"))
    return generate_portfolio_pack(input_dir=input_dir, output_dir=workspace / "deal_pack")


def _save_uploaded_inputs(workspace: Path, company_file, peer_file, precedent_file) -> Path:
    input_dir = workspace / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)
    for uploaded, name in [
        (company_file, "company_financials.csv"),
        (peer_file, "peer_comps.csv"),
        (precedent_file, "precedent_transactions.csv"),
    ]:
        (input_dir / name).write_bytes(uploaded.getvalue())
    return input_dir


def _render_dashboard(result: PortfolioPackResult) -> None:
    summary = result.summary
    implied = summary["adjusted_implied_enterprise_values"]
    ev_range = implied["range"]

    st.success("Portfolio deal pack generated from synthetic, public-safe sample data")

    metric_row = st.columns(6)
    metric_row[0].metric("Peers retained", summary["included_peer_count"])
    metric_row[1].metric("Peers excluded", summary["excluded_peer_count"])
    metric_row[2].metric("Precedents retained", summary["included_precedent_count"])
    metric_row[3].metric("Source status", summary["source_verification_status"])
    metric_row[4].metric("Unit status", summary["unit_normalization_status"])
    metric_row[5].metric("Artifacts", len(result.artifacts))

    st.subheader("Adjusted valuation range")
    st.markdown(f"## {ev_range['low']:,.2f} – {ev_range['high']:,.2f} {summary['target_unit']}")
    st.caption("Illustrative synthetic output only. Human review required. Not investment advice.")

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

    peer_tab, precedent_tab, json_tab, downloads_tab = st.tabs(
        ["Peer QA", "Precedent QA", "Summary JSON", "Downloads"]
    )
    with peer_tab:
        st.dataframe(summary["peer_review_rows"], use_container_width=True)
    with precedent_tab:
        st.dataframe(summary["precedent_review_rows"], use_container_width=True)
    with json_tab:
        st.json(summary)
    with downloads_tab:
        st.download_button(
            "Download complete portfolio deal pack",
            data=result.bundle_path.read_bytes(),
            file_name=result.bundle_path.name,
            mime="application/zip",
            use_container_width=True,
        )
        st.write("Generated artifacts")
        for artifact in result.artifacts:
            st.write(f"`{artifact.name}` · {artifact.stat().st_size:,} bytes")


st.title("DealForge AI — Investment Banking Analyst Copilot")
st.caption("Sanitized portfolio edition built by Mohit Bhatnagar · Synthetic data · Human review required")

st.markdown(
    """
This demo shows how structured financial inputs can be validated and converted into an adjusted valuation
summary, investment memo, diligence request list, Excel model, PowerPoint pitchbook, and portable ZIP pack.
The broader commercial product remains private under The Algosphere.
"""
)

with st.sidebar:
    st.header("Input mode")
    use_samples = st.checkbox("Use included synthetic sample data", value=True)
    company_file = st.file_uploader("company_financials.csv", type="csv", disabled=use_samples)
    peer_file = st.file_uploader("peer_comps.csv", type="csv", disabled=use_samples)
    precedent_file = st.file_uploader("precedent_transactions.csv", type="csv", disabled=use_samples)
    generate = st.button("Generate portfolio deal pack", type="primary", use_container_width=True)

if "portfolio_result" not in st.session_state:
    st.session_state["portfolio_result"] = _generate_pack(Path("sample_data"))

if generate:
    if use_samples:
        st.session_state["portfolio_result"] = _generate_pack(Path("sample_data"))
    else:
        if not all([company_file, peer_file, precedent_file]):
            st.error("Upload all three CSV files or use the included sample data.")
            st.stop()
        workspace = Path(tempfile.mkdtemp(prefix="dealforge_portfolio_uploads_"))
        input_dir = _save_uploaded_inputs(workspace, company_file, peer_file, precedent_file)
        st.session_state["portfolio_result"] = _generate_pack(input_dir)

_render_dashboard(st.session_state["portfolio_result"])

st.divider()
st.caption("Portfolio demonstration only. Human review required. Not investment advice.")
