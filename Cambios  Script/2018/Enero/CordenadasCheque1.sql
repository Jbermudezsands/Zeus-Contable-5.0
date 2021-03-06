/*
   lunes, 29 de enero de 201805:40:23 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2005
   Base de datos: SistemaContableBahingsa
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
	CaracteresConcepto nvarchar(50) NULL
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_CaracteresConcepto DEFAULT 0 FOR CaracteresConcepto
GO
COMMIT
select Has_Perms_By_Name(N'dbo.CordenadasCheque', 'Object', 'ALTER') as ALT_Per, Has_Perms_By_Name(N'dbo.CordenadasCheque', 'Object', 'VIEW DEFINITION') as View_def_Per, Has_Perms_By_Name(N'dbo.CordenadasCheque', 'Object', 'CONTROL') as Contr_Per 