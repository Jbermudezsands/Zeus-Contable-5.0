/*
   viernes, 08 de noviembre de 201901:32:04 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2005
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
ALTER TABLE dbo.IndiceTransaccion ADD
	Ajuste nvarchar(50) NULL
GO
COMMIT
