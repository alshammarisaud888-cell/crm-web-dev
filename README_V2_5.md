# Saudi Sensing CRM V2.5

## Fixes and New Workflow

### Meeting ID Label

- The Auto Meeting ID label is now displayed above the field.
- The label is no longer clipped by the TextField border.
- The Meeting ID remains automatic, unique and read-only.

### Send Opportunity to Quotation

A new button is available in every Opportunity row:

`Send to Quotation`

When clicked:

1. The CRM asks for confirmation.
2. A unique quotation request ID is generated automatically.
3. The request is linked to the selected Opportunity.
4. The Account is inherited from that Opportunity.
5. Status is set to `NEW REQUEST`.
6. Owner is set to `Proposal Engineer`.
7. The request appears at the top of the Quotations section.
8. The opportunity button changes to `Sent` while a pending request exists.

### Proposal Engineer Queue

The Quotations section now:

- Places NEW REQUEST records first.
- Shows a NEW indicator.
- Shows the linked Opportunity ID and Opportunity Name.
- Shows the inherited Account.
- Provides Edit and Delete actions.
- Allows the proposal engineer to open the request, add pricing and change status to Draft or Internal Review.

## Login

Username: admin
Password: admin123
