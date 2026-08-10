# Automated Operational & IT Risk Register

An automated Python solution designed to process, quantify, and report Operational and IT risks. This project ingests raw risk data, automatically calculates risk scores and severity metrics based on a 5 × 5 risk matrix, and generates an executive-ready Excel dashboard with visual analytics.

## Key Features

- **Automated Risk Scoring**: Calculates total Risk Score using the standard matrix formula:
  $$\text{Risk Score} = \text{Probability} \times \text{Impact}$$
- **Dynamic Severity Classification**: Automatically categorizes risks into 4 distinct tiers:
  - **Critical**: Score $\ge 15$
  - **High**: Score $10 - 14$
  - **Medium**: Score $5 - 9$
  - **Low**: Score $< 5$
- **Priority Ranking**: Automatically sorts the entire dataset by `Risk Score` and `Probability` in descending order, putting high-priority items at the top.
- **Executive Presentation**: Uses `XlsxWriter` to format the final Excel output with:
  - An **Executive Summary Dashboard** including KPI metric cards and an embedded doughnut chart.
  - A **Full Risk Register** sheet with color-coded severity indicators.
  - An isolated **Top Priority Open Risks** tab filtering for immediate mitigation actions (Critical/High & Open).

## Project Structure

```text
risk_automation/
├── risk_register_practice_data.xlsx  # Input dataset (raw risk list)
├── automated_risk_register.py        # Main Python automation script
├── Automated_Operational_Risk_...xlsx # Generated output Excel file
└── README.md                          # Project documentation
