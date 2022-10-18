
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 10/19/2022 00:52:14
-- Generated from EDMX file: C:\Users\Gamer1070\Desktop\rinaz\Template4438\Template4438\Safiullin_4438\Database\ExcelEntity.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [C:\USERS\GAMER1070\DOCUMENTS\DBDBDB.MDF];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[ExcelEntitySet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[ExcelEntitySet];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'ExcelEntitySet'
CREATE TABLE [dbo].[ExcelEntitySet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [ServiceName] nvarchar(max)  NOT NULL,
    [ServiceType] nvarchar(max)  NOT NULL,
    [ServiceCode] nvarchar(max)  NOT NULL,
    [ServicePrice] int  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'ExcelEntitySet'
ALTER TABLE [dbo].[ExcelEntitySet]
ADD CONSTRAINT [PK_ExcelEntitySet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------