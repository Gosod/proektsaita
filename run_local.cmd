@echo off
set JWT_SECRET=local-test-secret
set SPREADSHEET_ID=dummy
set ADMIN_IDS=699229724,924261386
set SERVE_STATIC=1
py api.py
pause
