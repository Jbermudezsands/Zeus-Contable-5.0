/*
   Sábado, 23 de Abril de 201110:35:00 a.m.
   Usuario: 
   Servidor: CONSULTOR\SQL2005
   Base de datos: SistemaContableBecas
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
ALTER TABLE dbo.Registros ADD
	VoucherNo nvarchar(50) NULL,
	TCambio float(53) NULL
GO
COMMIT
