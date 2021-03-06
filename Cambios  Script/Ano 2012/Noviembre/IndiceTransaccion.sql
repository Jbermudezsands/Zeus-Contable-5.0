/*
   sábado, 03 de noviembre de 201207:36:52 a.m.
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
ALTER TABLE dbo.IndiceTransaccion ADD
	ImprimeCheque bit NULL
GO
ALTER TABLE dbo.IndiceTransaccion ADD CONSTRAINT
	DF_IndiceTransaccion_ImprimeCheque DEFAULT 1 FOR ImprimeCheque
GO
COMMIT
