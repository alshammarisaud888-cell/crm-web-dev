# CRM V3.12 Azure Web Preview

This package is based on the supplied V3.12 Flet CRM rather than the temporary Flask demo.

## What this stage preserves

- Existing V3.12 Flet interface
- Existing SQLite CRM data as the initial seed database
- Accounts, contacts, leads, opportunities, quotations, dashboard, reports and analytics code
- Existing PDF / Excel generation logic
- Existing authentication and owner/workflow logic

## Changes made for Azure

- Flet starts as a browser-based web server.
- Azure uses `/home/crm` for persistent runtime data.
- The supplied database is copied into `/home/crm/data` on the first launch.
- Desktop-only Windows Explorer operations no longer crash the web server.
- Desktop mode remains available when `APP_ENV` is not set to `azure` or `web`.

## Azure App Service settings

Set:

- `APP_ENV` = `azure`
- `CRM_PERSIST_ROOT` = `/home/crm`
- `SCM_DO_BUILD_DURING_DEPLOYMENT` = `true`

Set the Startup Command to:

`bash startup.sh`

## Important limitations of this preview

This is a web-hosted preview of the current desktop application. It is **not yet the final multi-user production architecture**.

Before production use, the following need a dedicated web migration:

1. Replace SQLite with PostgreSQL for safe concurrent multi-user operation.
2. Move quotation/proposal attachments to Azure Blob Storage.
3. Add browser download endpoints for generated PDF/Excel reports.
4. Adapt Flet FilePicker upload behavior for browser uploads.
5. Add Microsoft Entra ID or another centralized identity provider.
6. Add server-side role enforcement, audit logs and production security controls.

The purpose of this package is to get the actual V3.12 interface running from Azure first, then migrate the data and document layers safely.
