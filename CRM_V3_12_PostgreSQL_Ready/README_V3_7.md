# Saudi Sensing CRM V3.7

## New Lead Fix

The New Lead form now opens correctly.

The issue was caused by updating the Company Name control before the dialog was mounted. The selected Account still fills Company Name automatically.

## Opportunity Table Alignment

- Opportunity ID and Account now have separate fixed-width columns.
- Added clear internal padding.
- Account was shifted slightly to the right.
- Long opportunity names have more space.
- Column spacing is controlled to avoid merged headings.

## Opportunity Workspace

Click the blue Opportunity ID to open a dedicated full-page workspace.

It includes:

### CRM Data
- Opportunity ID
- Account ID and Account Name
- Opportunity Name
- Project Type
- CRM Stage and Status
- Probability
- Gross and Weighted Values
- Expected Close Date
- Sales Owner
- Technical Owner
- Next Step
- Notes

### Original Excel Data
- Original Opportunity Reference
- Source Row and Created Date
- Customer Name from Source
- Account Basis
- Business Unit
- Source Project Type
- End User
- Industry
- Competitive status
- Probability Band
- Source Currency
- Original Stage
- Forecast GM %
- Forecast GM Value
- GM Value Basis
- Expected PO Year and Month
- Quarter
- Delivery Date
- Created By
- Assigned To
- Include in Forecast
- Must Win
- Suspended
- Opportunity Update
- Data Quality Flags
