-- ============================================================================
-- RMA Management System - SQL Server schema (single flat table)
-- Run once, as a user allowed to create databases/tables.
-- Dates are stored as 'YYYY-MM-DD' text and timestamps as 'YYYY-MM-DD HH:MM:SS'
-- text, because the application compares and slices them as strings.
-- Child data (repair entries, audit log, attachment paths) are JSON strings.
-- ============================================================================

IF DB_ID(N'RMA_DB') IS NULL
    CREATE DATABASE RMA_DB;
GO

USE RMA_DB;
GO

IF OBJECT_ID(N'dbo.rma', N'U') IS NULL
CREATE TABLE dbo.rma (
    id                      INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    rma_number              NVARCHAR(32)   NOT NULL,
    ncr_number              NVARCHAR(100)  NOT NULL DEFAULT N'N/A',
    project                 NVARCHAR(100)  NOT NULL,
    rma_category            NVARCHAR(50)   NOT NULL,
    model                   NVARCHAR(100)  NOT NULL,
    revision                NVARCHAR(50)   NOT NULL DEFAULT N'N/A',
    latest_revision         NVARCHAR(50)   NOT NULL DEFAULT N'N/A',
    serial_number           NVARCHAR(100)  NOT NULL,
    rework_notice           NVARCHAR(100)  NOT NULL DEFAULT N'N/A',
    received_date           NVARCHAR(10)   NOT NULL,
    received_by             NVARCHAR(100)  NOT NULL,
    failure_desc            NVARCHAR(MAX)  NOT NULL DEFAULT N'N/A',
    pre_visual              NVARCHAR(20)   NOT NULL DEFAULT N'Pending',
    pre_ict                 NVARCHAR(20)   NOT NULL DEFAULT N'Pending',
    pre_fct                 NVARCHAR(20)   NOT NULL DEFAULT N'Pending',
    post_visual             NVARCHAR(20)   NOT NULL DEFAULT N'N/A',
    post_ict                NVARCHAR(20)   NOT NULL DEFAULT N'N/A',
    post_fct                NVARCHAR(20)   NOT NULL DEFAULT N'N/A',
    failure_code            NVARCHAR(200)  NOT NULL DEFAULT N'N/A',
    shipment_date           NVARCHAR(10)   NOT NULL DEFAULT N'N/A',
    po_number               NVARCHAR(100)  NOT NULL DEFAULT N'N/A',
    shipment_remarks        NVARCHAR(MAX)  NOT NULL DEFAULT N'N/A',
    [status]                NVARCHAR(30)   NOT NULL DEFAULT N'Open',
    notes                   NVARCHAR(MAX)  NOT NULL DEFAULT N'N/A',
    cost_estimate           FLOAT          NOT NULL DEFAULT 0,
    pre_inspection_remarks  NVARCHAR(MAX)  NULL DEFAULT N'N/A',
    repair_entries          NVARCHAR(MAX)  NULL DEFAULT N'[]',   -- JSON array
    audit_log               NVARCHAR(MAX)  NULL DEFAULT N'[]',   -- JSON array
    attachment_paths        NVARCHAR(MAX)  NULL DEFAULT N'[]',   -- JSON array (local paths)
    created_at              NVARCHAR(19)   NULL,
    created_by              NVARCHAR(100)  NULL,
    updated_at              NVARCHAR(19)   NULL,
    updated_by              NVARCHAR(100)  NULL
);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_rma_rma_number')
    CREATE UNIQUE INDEX UX_rma_rma_number ON dbo.rma (rma_number);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_rma_serial_number')
    CREATE INDEX IX_rma_serial_number ON dbo.rma (serial_number);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_rma_status')
    CREATE INDEX IX_rma_status ON dbo.rma ([status]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_rma_received_date')
    CREATE INDEX IX_rma_received_date ON dbo.rma (received_date);
GO

-- Optional: least-privilege login for SQL authentication
-- CREATE LOGIN rma_user WITH PASSWORD = 'ChangeMe!123';
-- USE RMA_DB; CREATE USER rma_user FOR LOGIN rma_user;
-- ALTER ROLE db_datareader ADD MEMBER rma_user;
-- ALTER ROLE db_datawriter ADD MEMBER rma_user;
