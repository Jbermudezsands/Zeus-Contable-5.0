/*
   viernes, 06 de diciembre de 201902:37:16 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
   Base de datos: SistemaContableMakuCafe
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
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
ALTER TABLE dbo.Cuentas SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
