# Resume and Interview Snippets

## Resume project entry

**DealForge AI — Investment Banking Analyst Copilot**  
Built a Python and Streamlit finance automation system that converts structured company financials, comparable-company inputs, and precedent-transaction inputs into a 24-artifact investment-banking deal pack. The system validates valuation inputs, removes target-company self-inclusion, excludes off-sector transactions, handles NM EBITDA observations, generates adjusted valuation medians, and packages outputs into Excel, PowerPoint, memos, diligence lists, QA JSON artifacts, corrected CSV workpapers, and a checksum-verified ZIP bundle.

## Short resume bullets

- Built a public synthetic investment-banking analyst copilot using Python, Streamlit, OpenPyXL, and python-pptx.
- Generated a 24-artifact deal pack including Excel valuation workbook, PowerPoint pitchbook, memo, IC memo, diligence list, source-confidence report, valuation-audit report, corrected CSV workpapers, and manifest.
- Implemented valuation QA rules for target-company self-exclusion, sector mismatch, NM EBITDA handling, source-review flags, and transaction-unit normalization warnings.
- Added automated tests covering artifact generation, Office-file opening, bundle completeness, text hygiene, and JSON consistency.

## Interview answer: what problem did you solve?

Most finance portfolio projects stop at dashboards or notebooks. DealForge AI is built around the work product itself: the memo, model, deck, diligence list, source map, and validation reports must all agree. I designed the system so a user can see not only the valuation output but also which rows were included, which rows were removed, and why.

## Interview answer: why synthetic data?

The public repository is meant to demonstrate capability without exposing confidential data, customer documents, paid data sources, or private Algosphere product IP. Synthetic data lets the project remain safe and reproducible while still demonstrating realistic valuation QA patterns.

## Interview answer: what is technically interesting?

The interesting part is cross-artifact consistency. The same adjusted valuation summary drives Excel, PowerPoint, memo, diligence, JSON reports, and bundle manifest. Tests check that Office files open, the expected artifacts exist, and memo-style outputs do not leak Python null literals or omit human-review disclaimers.

## Interview answer: how would you productionize it?

The production Algosphere version would add real document ingestion, LLM/RAG over filings, page-level citations, saved deal workspaces, analyst-review workflows, authentication, organization-level isolation, audit logging, encryption, background jobs, and subscription controls. The public repo intentionally excludes those commercial and security layers.

## Interview answer: what finance judgment is encoded?

The public engine separates original valuation medians from adjusted medians. It removes the target company from the peer median, removes off-sector transactions from precedent medians, keeps NM EBITDA rows usable for EV/Revenue while excluding them from EV/EBITDA, and flags mixed transaction units instead of silently converting without proper assumptions.

## Disclaimer

Portfolio demonstration only. Human review required. Not investment advice.
