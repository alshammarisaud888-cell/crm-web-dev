# Saudi Sensing CRM V2.1 Fixed

## Fixes

- Fixed the blank white screen after login.
- Replaced Flet icons that may not exist in version 0.28.3.
- Added a protected startup error screen.
- Added `crm_startup_error.log` if startup fails.
- Added `RESET_AND_REINSTALL.bat` to remove an old incompatible environment.
- Preserved the preloaded 378 opportunity records and cleaned pipeline data.

## Recommended Start

1. Close the previous CRM window.
2. Extract this version into a completely new folder.
3. Do not copy the old `.venv` folder.
4. Run `run_app.bat`.
5. Login using:
   - Username: admin
   - Password: admin123

## If it still does not open

1. Run `RESET_AND_REINSTALL.bat`.
2. Run `run_app.bat` again.
3. If an error screen appears, send the text shown or the file:
   `crm_startup_error.log`
