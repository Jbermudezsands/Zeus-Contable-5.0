/*
   martes, 04 de agosto de 202004:06:36 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
   Base de datos: SistemaContableIncogasa
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
ALTER TABLE dbo.CatalogoActivoFijo ADD
	Localizacion nvarchar(50) NULL,
	FechaUltimaDepre smalldatetime NULL,
	ValorEstimadoMeses float(53) NULL,
	ValorRescate float(53) NULL,
	DepreciacionAcumulada float(53) NULL,
	CodEncargado nvarchar(50) NULL,
	CodCuenta nvarchar(50) NULL,
	NumeroMarbete nvarchar(50) NULL
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	DF_CatalogoActivoFijo_ValorEstimadoMeses DEFAULT 0 FOR ValorEstimadoMeses
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	DF_CatalogoActivoFijo_ValorRescate DEFAULT 0 FOR ValorRescate
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	DF_CatalogoActivoFijo_DepreciacionAcumulada DEFAULT 0 FOR DepreciacionAcumulada
GO
ALTER TABLE dbo.CatalogoActivoFijo SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
