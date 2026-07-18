from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from dealforge_portfolio import generate_portfolio_pack


st.set_page_config(page_title="DealForge AI Portfolio", page_icon="📊", layout="wide")

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

if generate:
    workspace = Path(tempfile.mkdtemp(prefix="dealforge_portfolio_"))
    if use_samples:
        input_dir = Path("sample_data")
    else:
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

    result = generate_portfolio_pack(input_dir=input_dir, output_dir=workspace / "deal_pack")
    summary = result.summary
    implied = summary["adjusted_implied_enterprise_values"]
    ev_range = implied["range"]

    st.success("Portfolio deal pack generated")
    metrics = st.columns(6)
    metrics[0].metric("Peers retained", summary["included_peer_count"])
    metrics[1].metric("Peers excluded", summary["excluded_peer_count"])
    metrics[2].metric("Precedents retained", summary["included_precedent_count"])
    metrics[3].metric("Source status", summary["source_verification_status"])
    metrics[4].metric("Unit status", summary["unit_normalization_status"])
    metrics[5].metric("Artifacts", len(result.artifacts))

    st.subheader("Adjusted valuation range")
    st.markdown(
        f"### {ev_range['low']:,.2f} – {ev_range['high']:,.2f} {summary['target_unit']}"
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
        for artifact in result.artifacts:
            st.write(f"`{artifact.name}` · {artifact.stat().st_size:,} bytes")

st.divider()
st.caption("Portfolio demonstration only. Human review required. Not investment advice.")
