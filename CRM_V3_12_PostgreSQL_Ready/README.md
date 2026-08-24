# Saudi Sensing CRM V1

A complete English CRM for Saudi Sensing, covering:

- Accounts and customers
- Contacts
- Leads
- Opportunities
- Pipeline stages and weighted forecast
- Quotations
- Proof of Concepts (PoCs)
- Meetings
- Activities and follow-ups
- Overdue action alerts
- Professional PDF reporting
- Excel export
- CSV-based data migration
- Database backup

## Default Login

- Username: admin
- Password: admin123

## Run

Double-click:

`run_app.bat`

## Existing Data Migration

1. Open the `migration_inbox` folder.
2. Fill the CSV templates, or export your current data into matching CSV files.
3. Keep the exact file names:
   - accounts.csv
   - contacts.csv
   - leads.csv
   - opportunities.csv
   - quotations.csv
   - pocs.csv
   - meetings.csv
   - activities.csv
4. Open the CRM.
5. Go to `Data Migration`.
6. Click `Import Migration Files`.

The application prevents duplicate imports using key fields such as account name, lead email, opportunity reference, quotation number and PoC reference.

## Important Files

Database:

`data/saudi_sensing_crm.db`

Reports:

`exports`

Migration files:

`migration_inbox`

## Business Scope

The CRM is tailored to Saudi Sensing activities such as instrumentation, automation, DCS, PLC, SCADA, cybersecurity, analyzers, metering, vibration monitoring, water, RO, PoCs, framework agreements, service agreements, localization and manufacturing initiatives.
