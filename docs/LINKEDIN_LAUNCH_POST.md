# LinkedIn Launch Post Draft

## Main post

I built **DealForge AI — Investment Banking Analyst Copilot**, a public portfolio project that turns structured finance inputs into a DealForge-style analyst work-product pack.

The project is designed around a realistic investment-banking workflow: selected financial metrics, comparable-company analysis, precedent-transaction analysis, valuation QA, source-confidence review, corrected workpapers, Excel model generation, PowerPoint pitchbook generation, investment memo, IC memo, diligence request list, consistency checks, and a checksum-verified ZIP bundle.

The public version is fully synthetic and safe to share, but intentionally demonstrates the visible workflow depth of a serious finance automation product.

What it generates:

- 24-artifact deal pack
- 11-sheet Excel valuation workbook
- 10-slide PowerPoint pitchbook
- Investment memo and IC memo
- Diligence request list
- Source-confidence and valuation-audit JSON artifacts
- Corrected peer and precedent CSV workpapers
- Portable ZIP bundle with manifest

The synthetic sample case intentionally includes common banker-review issues: the target company appears in its own peer set, one precedent is off-sector, several EBITDA observations are NM, and transaction values use mixed units. DealForge identifies these issues, records the exclusions, recalculates adjusted medians, and produces an adjusted EV range.

This project reflects the way I like to build: understand the business process, convert judgment-heavy steps into explicit controls, generate user-editable outputs, and test the workflow end to end.

Portfolio demonstration only. Human review required. Not investment advice.

#InvestmentBanking #Python #DataScience #FinTech #Streamlit #FinancialModeling #Automation #AIProducts #Analytics

## Short version

I built DealForge AI — a public synthetic Investment Banking Analyst Copilot that generates a 24-artifact deal pack with valuation QA, Excel model, PowerPoint pitchbook, IC memo, diligence list, source-confidence checks, corrected workpapers, and a checksum manifest.

The goal: show how a finance workflow can be turned into a tested, recruiter-visible software product.

## Comment follow-up

The public repository uses synthetic data only. The broader commercial product direction remains separate under The Algosphere.
