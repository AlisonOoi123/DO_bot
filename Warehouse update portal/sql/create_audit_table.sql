-- Run this once to set up the audit trail for the UVYDAY maintenance tool.
-- Records every change: who made it, when, which customer/route, and old -> new value per column.

IF OBJECT_ID('ENGSHENG.UVYDAY_AUDIT_LOG') IS NOT NULL
    DROP TABLE ENGSHENG.UVYDAY_AUDIT_LOG;

CREATE TABLE ENGSHENG.UVYDAY_AUDIT_LOG (
    LOG_ID        INT IDENTITY(1,1) PRIMARY KEY,
    CHANGED_AT    DATETIME2       NOT NULL DEFAULT SYSDATETIME(),
    CHANGED_BY    NVARCHAR(50)    NOT NULL,     -- username from the login form
    BPCNUM_0      NVARCHAR(30)    NOT NULL,
    BPAADD_0      NVARCHAR(30)    NOT NULL,
    CATEGORY      NVARCHAR(20)    NOT NULL,     -- 'DAY' / 'TIME' / 'LORRY' / 'OUTSOURCE' / 'ROUTE'
    COLUMN_NAME   NVARCHAR(30)    NOT NULL,     -- e.g. UVYDAY1_0, ZUVYDAY15_0, DRN_0
    OLD_VALUE     TINYINT         NULL,
    NEW_VALUE     TINYINT         NOT NULL
);

CREATE INDEX IX_UVYDAY_AUDIT_LOOKUP
    ON ENGSHENG.UVYDAY_AUDIT_LOG (BPCNUM_0, BPAADD_0, CHANGED_AT);
