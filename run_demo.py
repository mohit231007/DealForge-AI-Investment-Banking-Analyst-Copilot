from pathlib import Path

from dealforge_portfolio import generate_portfolio_pack


if __name__ == "__main__":
    result = generate_portfolio_pack(Path("sample_data"), Path("outputs/portfolio_demo"))
    ev_range = result.summary["adjusted_implied_enterprise_values"]["range"]
    unit = result.summary["target_unit"]
    print(f"Generated {len(result.artifacts)} artifacts in {result.output_dir}")
    print(f"Adjusted EV range: {ev_range['low']:,.2f} - {ev_range['high']:,.2f} {unit}")
    print(f"Bundle: {result.bundle_path}")
