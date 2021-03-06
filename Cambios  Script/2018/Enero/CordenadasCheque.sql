/*
   lunes, 29 de enero de 201812:50:50 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2005
   Base de datos: SistemaContableRevetsa
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar esta secuencia de comandos detalladamente antes de ejecutarla fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.CordenadasCheque ADD
	X14 nvarchar(50) NULL,
	Y14 nvarchar(50) NULL,
	X15 nvarchar(50) NULL,
	Y15 nvarchar(50) NULL,
	X16 nvarchar(50) NULL,
	Y16 nvarchar(50) NULL,
	X17 nvarchar(50) NULL,
	Y17 nvarchar(50) NULL,
	X18 nvarchar(50) NULL,
	Y18 nvarchar(50) NULL,
	X19 nvarchar(50) NULL,
	Y19 nvarchar(50) NULL,
	X20 nvarchar(50) NULL,
	Y20 nvarchar(50) NULL,
	X21 nvarchar(50) NULL,
	Y21 nvarchar(50) NULL,
	X22 nvarchar(50) NULL,
	Y22 nvarchar(50) NULL,
	X23 nvarchar(50) NULL,
	Y23 nvarchar(50) NULL,
	X24 nvarchar(50) NULL,
	Y24 nvarchar(50) NULL,
	X25 nvarchar(50) NULL,
	Y25 nvarchar(50) NULL
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X14 DEFAULT 0 FOR X14
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y14 DEFAULT 0 FOR Y14
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X15 DEFAULT 0 FOR X15
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y15 DEFAULT 0 FOR Y15
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X16 DEFAULT 0 FOR X16
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y16 DEFAULT 0 FOR Y16
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X17 DEFAULT 0 FOR X17
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y17 DEFAULT 0 FOR Y17
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X18 DEFAULT 0 FOR X18
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y18 DEFAULT 0 FOR Y18
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X19 DEFAULT 0 FOR X19
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y19 DEFAULT 0 FOR Y19
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X20 DEFAULT 0 FOR X20
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y20 DEFAULT 0 FOR Y20
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X21 DEFAULT 0 FOR X21
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y21 DEFAULT 0 FOR Y21
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X22 DEFAULT 0 FOR X22
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y22 DEFAULT 0 FOR Y22
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X23 DEFAULT 0 FOR X23
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y23 DEFAULT 0 FOR Y23
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X24 DEFAULT 0 FOR X24
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y24 DEFAULT 0 FOR Y24
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_X25 DEFAULT 0 FOR X25
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Y25 DEFAULT 0 FOR Y25
GO
COMMIT
select Has_Perms_By_Name(N'dbo.CordenadasCheque', 'Object', 'ALTER') as ALT_Per, Has_Perms_By_Name(N'dbo.CordenadasCheque', 'Object', 'VIEW DEFINITION') as View_def_Per, Has_Perms_By_Name(N'dbo.CordenadasCheque', 'Object', 'CONTROL') as Contr_Per 