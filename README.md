# Global Fashion Retail Sales 

## Overview
Interactive PySide6 desktop app to analyze customer LTV, cohorts, revenue structure, and run basic statistical tests.

- LTV factors by selected dimensions
- LTV cohort dynamics (6 months)
- Revenue structure comparison (pie charts)
- Statistical tests: Chi-square and T-test
- Executive Summary rendered from a local PDF

## Data
The app loads precomputed metrics from `customer_stats.csv` in the project root. Keep this file alongside the app.

## Requirements
- Python 3.10+
- See `requirements.txt` for Python dependencies

## Install
1. Create and activate a virtual environment (recommended)
2. Install dependencies:
   - `pip install -r requirements.txt`

## Run
- `python main.py`

## Executive Summary (PDF)
- Place a file named `Executive_summary.pdf` in the project root (next to `gui_app.py` and `main.py`).
- The app uses the native Qt PDF viewer (multi-page, fit-to-width).

## Data sourse link
https://www.kaggle.com/datasets/ricgomes/global-fashion-retail-stores-dataset

## Authors
- Darya Korenman 
- Ilia Usovich

