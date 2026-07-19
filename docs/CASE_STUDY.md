# Portfolio Case Study

## Problem

Investment-banking and investment-team workflows require more than a model or dashboard. Analysts must reconcile financial inputs, select comparable companies, review precedent transactions, document exclusions, triangulate valuation outputs, and communicate the result consistently across Excel, PowerPoint, memos, audit trails, source maps, and diligence lists.

## My approach

I designed DealForge AI as a work-product compiler rather than a generic finance chatbot.

The public portfolio-plus edition accepts three structured synthetic inputs:

1. selected company financial metrics;
2. comparable-company trading multiples;
3. precedent-transaction multiples and values.

It then creates a DealForge-style public-safe deal pack: source index, citation map, financial extract, comps analysis, precedent analysis, source confidence, valuation input audit, adjustment recommendations, corrected CSV workpapers, adjusted valuation summary, consistency checks, Excel workbook, PowerPoint pitchbook, investment memo, IC memo, diligence request list, manifest, and ZIP bundle.

## Key product decisions

### 1. Exclusions are visible

The engine records whether every peer and precedent is included or excluded and explains why. This avoids a black-box median calculation.

### 2. Revenue and EBITDA methods are separate

A row can remain in the EV/Revenue set while being excluded from EV/EBITDA because EBITDA is `NM` or not economically meaningful.

### 3. Mixed currencies are flagged, not guessed

The public edition does not silently convert transaction values without valuation dates and FX assumptions. It returns `NEEDS_NORMALIZATION`.

### 4. Source confidence is explicit

The source-confidence artifact tells reviewers that synthetic portfolio inputs are safe for demonstration but require replacement before any real banking use.

### 5. Outputs are consistent across artifacts

The same adjusted valuation summary feeds the memo, IC memo, Excel model, pitchbook, diligence request list, validation report, and downloadable bundle.

### 6. Public and commercial IP are separated

This repository is synthetic and public-safe. The private Algosphere edition can move much further with AI integrations, document intelligence, persistent workspaces, user accounts, approvals, security controls, and monetization.

## Results demonstrated by the sample case

The synthetic sample intentionally contains:

- the target company inside its own peer set;
- two peers with `NM` EBITDA;
- one off-sector precedent transaction;
- transaction values in INR crore, USD million, and USD billion;
- synthetic source notes that require source replacement.

The engine correctly:

- removes the target from the primary peer median;
- retains valid EV/Revenue observations while excluding `NM` EBITDA observations;
- removes the off-sector transaction;
- recalculates adjusted trading and transaction medians;
- produces an adjusted EV range of INR 45,000–80,000 crore;
- creates valuation audit warnings and adjustment recommendations;
- generates corrected peer and precedent workpapers;
- flags source and transaction-unit review requirements;
- packages the full output into a checksum-verified ZIP deal pack.

## Skills demonstrated

- investment-banking workflow understanding;
- valuation and comparable-company logic;
- source lineage and auditability;
- Python architecture and data validation;
- Excel automation with OpenPyXL;
- PowerPoint automation with python-pptx;
- Streamlit product design;
- automated QA and CI;
- documentation for technical and non-technical audiences;
- product strategy and IP separation.

## What I would contribute as an employee

This project reflects how I work:

1. understand the business process before coding;
2. convert judgment-heavy steps into explicit controls;
3. build outputs that real users can review and edit;
4. test failure cases, not only happy paths;
5. communicate limitations honestly;
6. think beyond a notebook toward a maintainable product.

## Disclaimer

This is a synthetic portfolio demonstration. It is not investment advice, a fairness opinion, or a certified valuation.
