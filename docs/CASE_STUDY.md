# Portfolio Case Study

## Problem

Investment-banking and investment-team workflows require more than a model or dashboard. Analysts must reconcile financial inputs, select comparable companies, review precedent transactions, document exclusions, triangulate valuation outputs, and communicate the result consistently across Excel, PowerPoint, memos, and diligence lists.

## My approach

I designed DealForge AI as a work-product compiler rather than a generic finance chatbot.

The portfolio edition accepts three structured inputs:

1. selected company financial metrics;
2. comparable-company trading multiples;
3. precedent-transaction multiples and values.

It then applies explicit review rules, produces original and adjusted medians, calculates directional implied enterprise values, and packages the results into analyst-friendly artifacts.

## Key product decisions

### 1. Exclusions are visible

The engine records whether every peer and precedent is included or excluded and explains why. This avoids a black-box median calculation.

### 2. Revenue and EBITDA methods are separate

A row can remain in the EV/Revenue set while being excluded from EV/EBITDA because EBITDA is `NM` or not economically meaningful.

### 3. Mixed currencies are flagged, not guessed

The public edition does not silently convert transaction values without valuation dates and FX assumptions. It returns `NEEDS_NORMALIZATION`.

### 4. Outputs are consistent

The same adjusted valuation summary feeds the memo, Excel model, pitchbook, diligence request list, JSON output, and downloadable bundle.

### 5. Public and commercial IP are separated

This repository is deliberately synthetic and compact. The private Algosphere edition is being developed as a broader product with workspace persistence, document ingestion, source lineage, user management, billing, and enterprise controls.

## Results demonstrated by the sample case

The synthetic sample intentionally contains:

- the target company inside its own peer set;
- two peers with `NM` EBITDA;
- one off-sector precedent transaction;
- transaction values in INR crore, USD million, and USD billion.

The engine correctly:

- removes the target from the primary peer median;
- retains valid EV/Revenue observations while excluding `NM` EBITDA observations;
- removes the off-sector transaction;
- recalculates adjusted trading and transaction medians;
- produces an adjusted EV range;
- flags source and transaction-unit review requirements.

## Skills demonstrated

- investment-banking workflow understanding;
- valuation and comparable-company logic;
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
