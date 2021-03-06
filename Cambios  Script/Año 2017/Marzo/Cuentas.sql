/*
   viernes, 31 de marzo de 201710:50:03 a.m.
   Usuario: sa
   Servidor: SERVER\SQL2005
   Base de datos: SistemaContable
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
ALTER TABLE dbo.Cuentas ADD
	CentroCostos bit NULL
GO
ALTER TABLE dbo.Cuentas ADD CONSTRAINT
	DF_Cuentas_CentroCostos DEFAULT 0 FOR CentroCostos
GO
COMMIT
