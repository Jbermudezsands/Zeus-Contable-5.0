/*
   sábado, 19 de enero de 201308:22:40 a.m.
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
COMMIT
BEGIN TRANSACTION
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.DetalleEgreso ADD CONSTRAINT
	FK_DetalleEgreso_IndiceEgreso FOREIGN KEY
	(
	NumeroEgreso
	) REFERENCES dbo.IndiceEgreso
	(
	NumeroEgreso
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
ALTER TABLE dbo.DetalleEgreso ADD CONSTRAINT
	FK_DetalleEgreso_Cuentas FOREIGN KEY
	(
	CodCuentas
	) REFERENCES dbo.Cuentas
	(
	CodCuentas
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
