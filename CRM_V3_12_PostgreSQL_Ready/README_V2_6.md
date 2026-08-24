# Saudi Sensing CRM V2.6

## Owners Master

A new Owners section replaces the PoCs menu.

Owner records contain:

- Owner ID
- Owner Name
- Role
- Department
- Email
- Mobile
- Quotation assignment eligibility
- Status
- Notes

Quotation Owner is now a dropdown populated only from active owners marked as eligible for quotation assignments.

## Sequential IDs

The CRM now uses automatic sequential numbers:

- Accounts: ACC-00001
- Contacts: CON-00001
- Leads: LD-00001
- Opportunities: OPP-00001
- Quotations: QTN-00001
- Meetings: MTG-00001
- Activities: ACT-00001
- Owners: OWN-00001

All new records receive the next number automatically.

Existing records are renumbered by database order. Original imported opportunity references remain available in Migrated Pipeline Details.

## Important

The PoCs data table remains in the database for compatibility, but the PoCs menu is removed as requested.
