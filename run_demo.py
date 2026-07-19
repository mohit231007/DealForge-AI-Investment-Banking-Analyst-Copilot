from __future__ import annotations

from datetime import datetime
from pathlib import Path

from dealforge_portfolio import generate_portfolio_pack


DEFAULT_OUTPUT_DIR = Path("outputs/portfolio_demo")


def _has_locked_file(path: Path) -> bool:
    """Return True when an existing output file is currently locked by Excel/PowerPoint.

    On Windows, Office keeps generated .xlsx/.pptx files locked while they are open.
    The demo should still be runnable, so it falls back to a timestamped folder.
    """

    if not path.exists():
        return False
    for candidate in path.iterdir():
        if not candidate.is_file():
            continue
        try:
            with candidate.open("ab"):
                pass
        except PermissionError:
            return True
    return False


def _resolve_output_dir() -> Path:
    if _has_locked_file(DEFAULT_OUTPUT_DIR):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = Path("outputs") / f"portfolio_demo_{stamp}"
        print(
            "Default output folder has an Office file open. "
            f"Writing this run to {fallback} instead."
        )
        return fallback
    return DEFAULT_OUTPUT_DIR


if __name__ == "__main__":
    result = generate_portfolio_pack(Path("sample_data"), _resolve_output_dir())
    ev_range = result.summary["adjusted_implied_enterprise_values"]["range"]
    unit = result.summary["target_unit"]
    print(f"Generated {len(result.artifacts)} artifacts in {result.output_dir}")
    print(f"Adjusted EV range: {ev_range['low']:,.2f} - {ev_range['high']:,.2f} {unit}")
    print(f"Bundle: {result.bundle_path}")
