-- Run this once in SSMS (as sa or another admin) to create a narrowly-scoped
-- login for the UVYDAY maintenance web app - do NOT point the app at 'sa'.
-- This login can only read/write the two tables it actually needs.

USE [x3v12];
GO

-- 1. Create the SQL login (change the password!)
CREATE LOGIN uvyday_app_user WITH PASSWORD = 'ChangeThisToAStrongPassword!';
GO

-- 2. Create a database user mapped to that login
CREATE USER uvyday_app_user FOR LOGIN uvyday_app_user;
GO

-- 3. Grant only what the app needs: read/write on BPDLVCUST, read/write on the audit log
GRANT SELECT, UPDATE ON ENGSHENG.BPDLVCUST TO uvyday_app_user;
GRANT SELECT, INSERT ON ENGSHENG.UVYDAY_AUDIT_LOG TO uvyday_app_user;
GO

-- That's it - no sysadmin role, no db_owner, no access to any other table.
