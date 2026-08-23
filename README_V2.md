# Saudi Sensing CRM V2 - Opportunity-Based Edition

This version is built around the uploaded Saudi Sensing Opportunity Pipeline.

## Preloaded Data

- 378 opportunities
- Accounts generated from Customer Name, or End User when Customer Name was missing
- Original source fields preserved in `Migrated Pipeline Details`
- CRM stages, probability percentages and statuses mapped from the source file
- Data quality flags included for review

## Important Data Decisions

- No currency conversion was performed because the source currency was not consistently confirmed.
- Probability mapping:
  - A = 90%
  - B = 70%
  - C = 40%
  - D = 20%
- Stage mapping:
  - Budgetary -> Qualification
  - Firm -> Proposal
  - Won -> Awarded / Won
  - Lost and Canceled -> Lost
  - On-Hold -> On Hold
- Missing GM values were calculated only when Gross and GM % were available.
- Material GM discrepancies are flagged and not silently corrected.

## Login

- Username: admin
- Password: admin123

## Run

Double-click `run_app.bat`.

## Cleaned Source Workbook

The package includes:

`Saudi_Sensing_Opportunity_Pipeline_Cleaned_for_CRM.xlsx`

It contains:
- Dashboard
- Cleaned Opportunities
- Accounts
- CRM Opportunities Import
- Field Mapping
