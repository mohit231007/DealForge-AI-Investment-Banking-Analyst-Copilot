# Interview Guide

Use this document to explain the project clearly in interviews without overselling it.

## 30-second explanation

> I built DealForge AI to automate a repeatable investment-banking workflow rather than create another finance chatbot. The public portfolio edition takes structured company financials, peer comps, and precedent transactions; validates them; removes unsuitable observations; calculates original and adjusted valuation outputs; and generates an investment memo, diligence list, Excel model, PowerPoint deck, and audit-ready ZIP pack.

## What was technically difficult

### Financial and valuation QA

The hard part was not calculating a median. It was defining when a row should be included in EV/Revenue but excluded from EV/EBITDA, detecting target-company self-inclusion, identifying sector mismatch, and avoiding unsafe automatic assumptions.

### Cross-artifact consistency

Every output should use the same selected values and adjusted valuation layer. A memo, workbook, deck, and JSON file that disagree are more dangerous than no output.

### Office-file reliability

Generated Excel and PowerPoint files must open in standard desktop applications. The project uses OpenPyXL and python-pptx rather than fragile hand-written Office XML.

### Product boundary

I deliberately created a public synthetic edition for recruiters and kept the larger Algosphere product private. This demonstrates commercial and IP awareness, not just coding.

## Questions I expect

### Is this investment advice?

No. It is a workflow and QA demonstration. All outputs are starter work products requiring human review.

### Why not use an LLM for every step?

Valuation calculations and inclusion rules should be deterministic and testable. Generative AI is more appropriate for controlled commentary after core calculations and source checks are established.

### Why use synthetic data?

It makes the repository safe to share, reproducible, and free of confidential or licensed data while preserving the workflow challenges.

### What would you build next in a commercial version?

- authenticated deal workspaces;
- document ingestion and page-level citations;
- saved versions and reviewer comments;
- market-data and filing connectors;
- billing and entitlements;
- role-based access and enterprise audit logs;
- cloud and private-deployment options.

## Resume bullets

- Built a Python/Streamlit investment-banking analyst copilot that validates comparable-company and precedent-transaction inputs and generates adjusted valuation outputs, Excel models, PowerPoint pitchbooks, memos, diligence lists, and ZIP deal packs.
- Designed deterministic valuation QA controls for target self-inclusion, sector mismatch, outlier multiples, `NM` EBITDA, source verification, and transaction-unit normalization.
- Automated editable Office work products using OpenPyXL and python-pptx, with CI tests covering valuation calculations, file integrity, text hygiene, and bundle completeness.
- Separated a recruiter-facing synthetic portfolio edition from a private commercial product roadmap under The Algosphere.

## Honest limitations to mention

- The public edition uses structured CSVs rather than full document intelligence.
- The sample data is fictional.
- Market and transaction values are not live.
- Final valuation judgment remains with a qualified human analyst.
