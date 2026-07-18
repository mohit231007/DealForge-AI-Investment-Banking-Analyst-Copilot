# Technical Architecture

## Objective

The public portfolio edition demonstrates a deterministic analyst workflow without publishing the private Algosphere product core or using confidential company data.

## Components

```text
app.py
  └── recruiter-facing Streamlit interface

run_demo.py
  └── command-line entry point

dealforge_portfolio/engine.py
  ├── CSV loading and numeric normalization
  ├── peer-set QA
  ├── precedent-transaction QA
  ├── original and adjusted median calculation
  ├── implied enterprise-value triangulation
  ├── memo and diligence generation
  ├── OpenPyXL workbook generation
  ├── python-pptx deck generation
  └── manifest and ZIP packaging

sample_data/
  └── synthetic financial, peer, and transaction inputs

tests/
  └── valuation logic, Office-file, text-hygiene, and bundle QA
```

## QA logic

### Comparable companies

A peer is excluded from the primary EV/Revenue set when it is:

- the target company itself;
- from a different market;
- from a non-overlapping sector;
- missing EV/Revenue; or
- outside the 0–25x review range.

EV/EBITDA is independently excluded when it is `NM`, non-positive, or above 100x.

### Precedent transactions

The same market, sector, and multiple checks are applied to transactions. Mixed transaction-value units trigger a normalization warning rather than an unsafe automatic currency conversion.

## Export layer

The engine produces human-readable Markdown and JSON plus repair-safe Office files:

- Excel is generated using OpenPyXL;
- PowerPoint is generated using python-pptx;
- SHA-256 checksums are recorded in the manifest;
- all artifacts are packaged into one ZIP file.

## Design principles

1. **Deterministic before generative** — calculations should be reproducible.
2. **Visible exclusions** — rows are not silently dropped.
3. **Human review by default** — outputs are starter work products.
4. **Synthetic public data** — no confidential/private inputs are committed.
5. **Public/private boundary** — commercial features remain in the private Algosphere repository.
