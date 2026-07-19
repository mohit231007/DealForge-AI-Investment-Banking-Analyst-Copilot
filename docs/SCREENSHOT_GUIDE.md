# Screenshot Capture Guide

## Goal

Capture recruiter-ready visuals for the README, GitHub profile, LinkedIn post, and interview walkthrough.

Use only the included synthetic sample data. Do not show private directories, private repositories, real company documents, browser history, API keys, or local personal information.

## Recommended browser setup

- Use the local Streamlit app at `http://localhost:8501`.
- Collapse the sidebar for the main hero screenshot.
- Zoom browser to 90% or 100% depending on screen size.
- Use dark theme for consistency.
- Close unrelated browser tabs before final screenshots.

## Required screenshots

Save screenshots into an `assets/` folder using these names:

```text
assets/01_streamlit_dashboard.png
assets/02_valuation_qa.png
assets/03_source_confidence.png
assets/04_comps_analysis.png
assets/05_precedent_transactions.png
assets/06_downloads_artifacts.png
assets/07_excel_model.png
assets/08_pitchbook.png
```

## Shot list

### 1. Streamlit dashboard

Show the hero, banker takeaway, 24 artifacts card, adjusted EV range card, QA status, peer-set decision, precedent decision, and source/unit review badges.

### 2. Valuation QA

Open the Valuation QA tab and show the audit checks plus adjustment recommendations.

### 3. Source confidence

Open the Source Confidence tab and show source type, confidence score, source review status, and review notes.

### 4. Comps analysis

Open the Comps Analysis tab and show included/excluded peers, especially the target-company exclusion row.

### 5. Precedent transactions

Open the Precedent Transactions tab and show the off-sector transaction exclusion.

### 6. Downloads / artifacts

Open the Downloads tab and show the ZIP button plus the generated 24-artifact list.

### 7. Excel model

Open `outputs/portfolio_demo/10_valuation_model.xlsx` and capture the Adjusted Valuation QA sheet or the Source Index sheet.

### 8. Pitchbook

Open `outputs/portfolio_demo/11_pitchbook.pptx` and capture the Adjusted Valuation QA slide or title slide.

## Local commands

```powershell
cd C:\Users\mohit\DealForge-AI-Investment-Banking-Analyst-Copilot
python -m pytest
python run_demo.py
streamlit run app.py
```

## README insertion target

After assets are committed, add a visual gallery near the top of `README.md`, directly after the badge section.

Suggested markdown:

```markdown
## Visual walkthrough

![DealForge dashboard](assets/01_streamlit_dashboard.png)
![Valuation QA](assets/02_valuation_qa.png)
![Source confidence](assets/03_source_confidence.png)
![Excel model](assets/07_excel_model.png)
![Pitchbook](assets/08_pitchbook.png)
```

## Final hygiene check

Before publishing screenshots, confirm:

- no private repo name is visible except the public repo;
- no local personal files are visible;
- no unrelated browser tabs are visible;
- no real customer/company confidential data is visible;
- disclaimer language remains visible in at least one screenshot.
