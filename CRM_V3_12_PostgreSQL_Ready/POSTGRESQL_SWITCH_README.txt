CRM V3.12 PostgreSQL-ready build

Safety design:
- Default remains SQLite.
- PostgreSQL is used only when the Azure App Setting
  DATABASE_BACKEND=postgres
  is added.
- Existing PGHOST, PGPORT, PGDATABASE, PGUSER and PGPASSWORD settings are used.
- The existing SQLite database remains untouched, so rollback is simply:
  DATABASE_BACKEND=sqlite

Prerequisite:
- Run migrate_sqlite_to_postgres.py successfully before switching.
- Verify row counts in PostgreSQL.

Switch procedure:
1. Deploy this build.
2. Confirm deployment succeeds while DATABASE_BACKEND is still absent/sqlite.
3. In Azure App Service > Settings > Environment variables add:
   DATABASE_BACKEND = postgres
4. Apply/restart.
5. Test login, Dashboard, Accounts, Opportunities, Quotations, and one test write.
6. If anything fails, set DATABASE_BACKEND=sqlite and restart.

Important:
- Change the PostgreSQL password because it was exposed in a screenshot.
- Update PGPASSWORD in the App Service immediately after changing it.
