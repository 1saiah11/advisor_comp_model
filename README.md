# Advisor Compensation Modeling Automation

## Project Overview

This project automates the calculation, reporting, and analytics of advisor compensation for a wealth management firm. It was designed to handle complex commission structures, splits, management/SMA fees, and advisor-specific reporting using both Excel and Python.

## Background

Wealth management firms often face the challenge of accurately and efficiently calculating advisor compensation, especially when handling multiple commission structures, split rules, and dynamic reporting needs. Manual calculations can lead to errors, inefficiencies, and lack of transparency. This project was developed to automate and streamline the entire advisor compensation process. THe main focus was ensuring accuracy, scalability, and actionable reporting for stakeholders.

---

## Key Challenge

**Reconciling Business Logic with Code Automation**  
Translating nuanced compensation rules (commission splits, negative charge adjustments, etc) into both Excel formulas and Python logic required several iterations to ensure every edge case was handled correctly and outputs matched across tools.

---

## Key Features

- **Automated Data ETL:** Import, clean, and transform source data from Excel files using Python (pandas) and Excel Power Query.
- **Dynamic Calculations:** Apply bucket rates (Service, Sign, Intro), handle custom split logic, and compute commissions with minimal manual oversight.
- **Advisor-Specific Reports:** Automatically generate detailed payout statements for each advisor.
- **Summary Analytics:** Calculate min, max, average, variance, and additional payout distribution metrics for management insight.
- **Scalable & Transparent:** Model is easily extendable for new advisors, rules, or data sources.

---

## Technologies Used

- Python (`pandas`, `xlsxwriter`)
- Excel (Power Query, Pivot Tables, Dynamic Arrays)
- Jupyter Notebook

---

## Sample Files

- `Advisor_Billing.xlsx` – Example input data with management, SMA, and charges sheets
- `Advisor_Metrics.xlsx` – Example input for advisor rates and splits
- `Payout_<Advisor>_Final.xlsx` – Sample automated payout report for an advisor

---

## How to Use

1. Place your input data files in the working directory.
2. Run the Jupyter notebook or Python script.
3. Automated reports for each advisor will be generated in the output directory.
4. Review summary analytics or pivot tables for further insights.
