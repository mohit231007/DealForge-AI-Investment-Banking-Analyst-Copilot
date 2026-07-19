# DealForge AI Portfolio-Plus v1.0 Release Notes

## Release goal

This release positions DealForge AI as a banker-grade public portfolio project: strong enough to show investment-banking workflow depth to recruiters and hiring managers, while remaining fully synthetic and safe for public GitHub.

## What changed

The public repository now generates a DealForge-style portfolio-plus deal pack with 24 artifacts:

```text
01_company_profile.md
02_investment_memo.md
03_source_index.json
04_synthetic_document_chunks.jsonl
05_due_diligence_red_flags.md
06_audit_trail.md
07_citation_map.json
08_source_map.json
09_validation_report.json
10_valuation_model.xlsx
11_pitchbook.pptx
12_financial_extract.json
13_comps_analysis.json
14_precedent_transactions.json
15_consistency_checks.json
16_investment_committee_memo.md
17_due_diligence_request_list.md
18_source_confidence_ladder.json
19_valuation_input_audit.json
20_valuation_adjustment_recommendations.json
21_corrected_peer_comps.csv
22_corrected_precedent_transactions.csv
23_adjusted_valuation_summary.json
00_deal_pack_bundle_manifest.json
```

## Demonstrated workflow

The synthetic case demonstrates:

- target-company self-inclusion detection;
- off-sector precedent exclusion;
- NM EBITDA handling;
- original versus adjusted valuation medians;
- adjusted implied enterprise value range;
- source-confidence review flags;
- unit-normalization review flags;
- corrected peer and precedent workpapers;
- cross-artifact consistency checks;
- Excel and PowerPoint export automation;
- checksum-based bundle packaging.

## Validated local result

```text
4 tests passed
24 generated artifacts
11-sheet Excel workbook
10-slide PowerPoint pitchbook
Adjusted EV range: INR 45,000–80,000 crore
```

## Public safety boundary

This release uses fictional company, peer, and transaction data. It does not include confidential documents, customer data, paid data connectors, private Algosphere commercial orchestration, or investment advice.

## Release checklist before tagging

- [ ] Clean clone works.
- [ ] `python -m pytest` passes.
- [ ] `python run_demo.py` generates 24 artifacts.
- [ ] `streamlit run app.py` opens the portfolio-plus control tower.
- [ ] Excel workbook opens without repair warnings.
- [ ] PowerPoint opens with 10 slides.
- [ ] README screenshots are added.
- [ ] LinkedIn launch post is published.
- [ ] GitHub repo is pinned on the profile.

## Disclaimer

Synthetic portfolio demonstration only. Human review required. Not investment advice.
