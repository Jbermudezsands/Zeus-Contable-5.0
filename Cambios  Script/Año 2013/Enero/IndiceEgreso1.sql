/*
   sábado, 19 de enero de 201308:53:01 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaContableIpemsa
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
ALTER TABLE dbo.IndiceEgreso ADD
	Activo bit NULL,
	ActivoCheque bit NULL
GO
ALTER TABLE dbo.IndiceEgreso ADD CONSTRAINT
	DF_IndiceEgreso_Activo DEFAULT 0 FOR Activo
GO
ALTER TABLE dbo.IndiceEgreso ADD CONSTRAINT
	DF_IndiceEgreso_ActivoCheque DEFAULT 0 FOR ActivoCheque
GO
COMMIT
