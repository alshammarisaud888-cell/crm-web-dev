# Saudi Sensing CRM V2.4

## Main Improvements

- Saudi Sensing logo added to login, sidebar and executive PDF report.
- Clearer meeting field labels and improved meeting layout.
- Edit and Delete actions added to:
  - Accounts
  - Contacts
  - Leads
  - Opportunities
  - Quotations
  - PoCs
  - Meetings
  - Activities
- Every new Lead, Opportunity, Quotation, PoC, Meeting and Activity receives a visible auto-generated unique ID.
- Unique IDs are read-only.
- Quotations must be linked to an Opportunity.
- Account is removed from manual quotation entry.
- The quotation account is inherited automatically from the selected Opportunity.
- Meetings may be linked to an Account only or to a specific Opportunity.
- Selecting an Opportunity in Meetings updates the related Account automatically.

## Relationship Model

Account
  -> Contacts
  -> Opportunities
       -> Quotations
       -> PoCs
       -> Meetings
       -> Activities

Leads remain separate until qualified and converted into Accounts and Opportunities.

## Login

Username: admin
Password: admin123
