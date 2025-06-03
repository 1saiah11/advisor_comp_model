# Advisor Compensation Modeling Automation

## Background

Wealth management firms face challenges in calculating advisor compensation accurately and efficiently. This is especially true when dealing with multiple commission structures, complex split rules, and the need for dynamic reporting. Manual processes often introduce errors, inefficiencies, and a lack of transparency. This project was developed to automate and streamline the advisor compensation process, ensuring accuracy, scalability, and actionable reporting for stakeholders.

## Key Challenge

Reconciling business logic with code automation was essential. Translating nuanced compensation rules, especially those involving commission splits and negative charge adjustments, into both Excel formulas and Python code required several iterations. This process was necessary to make sure every edge case was handled correctly and that outputs matched across tools.

## Dual Approach

This project includes two fully functional automation methods:

### Power Query and Excel Template

The Power Query and Excel solution provides a user-friendly, visually polished template for business users. This approach is ideal for teams who prefer working within the familiar Excel interface. The report layouts, summary statistics, and dynamic filtering allow for quick onboarding and easy adoption.

![Screenshot 2025-06-03 143712](https://github.com/user-attachments/assets/21645681-56ae-42a4-9190-0c984ef555dd)


### Python Automation

The Python implementation uses pandas and xlsxwriter to automate data transformation, commission calculations, and advisor-specific report generation. This approach is robust and scalable, making it suitable for large datasets or situations where automation and reproducibility are key.

![Screenshot 2025-06-03 144330](https://github.com/user-attachments/assets/f03ec018-cd1c-4419-b700-842e5f9f46f3)
*Note: Full code is available in both the Jupyter Notebook (.ipynb) and an exported HTML version for easy viewing, included in this repository*
## Key Features

* Automated data import, cleaning, and transformation using Python and Power Query
* Dynamic commission and split calculations that reflect real-world business logic
* Advisor-specific payout summaries with breakdowns by fee type and split scenario
* Summary analytics including minimum, maximum, average, and variance of payouts
* Visual and export-ready Excel reports for both technical and non-technical audiences

## Technologies Used

* Python (pandas, xlsxwriter)
* Excel (Power Query, Pivot Tables)
* Jupyter Notebook

## How to Use

1. Place your input data files in the working directory.
2. Run the Jupyter notebook or Python script for automated report generation, or open the Excel template and refresh queries for manual/visual analysis.
3. Advisor-specific reports and summary statistics will be available in the output directory or Excel workbook.
4. Review the summary analytics and use the included dynamic filtering tools for deeper insight.

## Sample Files

* `Advisor_Billing.xlsx` — Example input data with management, SMA, and charges sheets
* `Advisor_Metrics.xlsx` — Advisor rates and splits
* `Payout_<Advisor>_Final.xlsx` — Example Python-generated payout report
* Excel Power Query template file

## Screenshots

### Power Query Dashboard
![Screenshot 2025-06-01 220353](https://github.com/user-attachments/assets/4199ca0c-9f79-4901-9557-2a63299dd1aa)


### Python Generated Excel Template
![image](https://github.com/user-attachments/assets/b64d081a-1c1c-400a-b381-c54ae5944f28)




