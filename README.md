# datagradsurvey

This repository now includes an automated analysis pipeline for the 2023 graduate exit survey workbook.

## What gets generated

Running the analysis creates `outputs/2023/` with:

- `course_ranking.csv` — full rank-ordered course results.
- `ranking.svg` — chart of the top-ranked courses.
- `README.md` — human-readable summary and top-5 table.
- `metadata.json` — run metadata (sheet name, response count, etc.).

## Run locally

```bash
python scripts/analyze_exit_survey.py
```

## Run in GitHub Actions

Workflow file: `.github/workflows/analyze-survey.yml`

- Triggers on:
  - manual run (`workflow_dispatch`)
  - pushes that change the workbook, analysis script, or workflow.
- Uploads `outputs/` as an artifact named **`survey-analysis-outputs`** for download.
