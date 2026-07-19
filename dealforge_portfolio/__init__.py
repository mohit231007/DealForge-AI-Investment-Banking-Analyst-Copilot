"""Public portfolio edition of DealForge AI."""

from __future__ import annotations

from pathlib import Path

from dealforge_portfolio.engine import PortfolioPackResult
from dealforge_portfolio.engine import generate_portfolio_pack as _generate_engine_pack
from dealforge_portfolio.premium_pitchbook import upgrade_pitchbook_and_bundle


def generate_portfolio_pack(
    input_dir: str | Path = "sample_data",
    output_dir: str | Path = "outputs/portfolio_demo",
) -> PortfolioPackResult:
    """Generate the portfolio-plus pack and upgrade the pitchbook presentation layer."""

    result = _generate_engine_pack(input_dir=input_dir, output_dir=output_dir)
    upgrade_pitchbook_and_bundle(result)
    return result


__all__ = ["generate_portfolio_pack"]
