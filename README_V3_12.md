# Saudi Sensing CRM V3.12

## PDF Export Repair

The Opportunity Workspace and Quotation Workspace PDF functions were missing from the previous build while the buttons still referenced them.

This version adds complete PDF export functions and reconnects both buttons.

## Opportunity PDF

Includes:

- CRM opportunity details
- Account information
- Original Excel data
- Values and probability
- Forecast gross margin
- Expected PO and delivery information
- Opportunity update
- CRM notes
- Data quality flags

## Quotation PDF

Includes:

- Quotation and linked opportunity details
- Commercial calculations
- Cost price
- VAT and total including VAT
- Gross margin percentage and value
- Workflow dates
- Approval status
- Assignment notes
- Proposal notes
- Approval comments
- Uploaded document register

## Export Confirmation

After successful export, the program shows:

- Exact PDF location
- Open PDF button
- Open Folder button

Exports are saved in the application's `exports` folder.

If an error occurs, a detailed log is created:

- `opportunity_pdf_export_error.log`
- `quotation_pdf_export_error.log`
