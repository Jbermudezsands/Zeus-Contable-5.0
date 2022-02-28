/*
   miércoles, 23 de febrero de 202215:50:24
   Usuario: 
   Servidor: JUANBERMUDEZ
   Base de datos: SistemaFacturacionMulukuku
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
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT FK_DetalleNomina_Productor
GO
ALTER TABLE dbo.Productor SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT FK_DetalleNomina_Nomina
GO
ALTER TABLE dbo.Nomina SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_TipoProductor
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Lunes
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Martes
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Miercoles
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Jueves
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Viernes
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Sabado
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Domingo
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Total
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioVenta
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_TotalIngresos
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Trazabilidad
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_OtrasDeducciones
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_Bolsa
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioLunes
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioMartes
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioMiercoles
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioJueves
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioViernes
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioSabado
GO
ALTER TABLE dbo.Detalle_Nomina
	DROP CONSTRAINT DF_Detalle_Nomina_PrecioDomingo
GO
CREATE TABLE dbo.Tmp_Detalle_Nomina
	(
	NumNomina nvarchar(50) NOT NULL,
	CodProductor nvarchar(50) NOT NULL,
	TipoProductor nvarchar(50) NOT NULL,
	Roc1 nvarchar(50) NULL,
	Lunes float(53) NULL,
	Roc2 nvarchar(50) NULL,
	Martes float(53) NULL,
	Roc3 nvarchar(50) NULL,
	Miercoles float(53) NULL,
	Roc4 nvarchar(50) NULL,
	Jueves float(53) NULL,
	Roc5 nvarchar(50) NULL,
	Viernes float(53) NULL,
	Roc6 nvarchar(50) NULL,
	Sabado float(53) NULL,
	Roc7 nvarchar(50) NULL,
	Domingo float(53) NULL,
	Total float(53) NULL,
	PrecioVenta float(53) NULL,
	TotalIngresos decimal(18, 2) NULL,
	IR decimal(18, 2) NULL,
	DeduccionPolicia decimal(18, 2) NULL,
	Anticipo decimal(18, 2) NULL,
	DeduccionTransporte decimal(18, 2) NULL,
	Pulperia decimal(18, 2) NULL,
	Inseminacion decimal(18, 2) NULL,
	ProductosVeterinarios decimal(18, 2) NULL,
	Trazabilidad decimal(18, 2) NULL,
	OtrasDeducciones decimal(18, 2) NULL,
	Nombres nvarchar(250) NULL,
	Bolsa decimal(18, 2) NULL,
	PrecioLunes float(53) NULL,
	PrecioMartes float(53) NULL,
	PrecioMiercoles float(53) NULL,
	PrecioJueves float(53) NULL,
	PrecioViernes float(53) NULL,
	PrecioSabado float(53) NULL,
	PrecioDomingo float(53) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina SET (LOCK_ESCALATION = TABLE)
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_TipoProductor DEFAULT ((0)) FOR TipoProductor
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Lunes DEFAULT ((0)) FOR Lunes
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Martes DEFAULT ((0)) FOR Martes
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Miercoles DEFAULT ((0)) FOR Miercoles
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Jueves DEFAULT ((0)) FOR Jueves
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Viernes DEFAULT ((0)) FOR Viernes
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Sabado DEFAULT ((0)) FOR Sabado
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Domingo DEFAULT ((0)) FOR Domingo
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Total DEFAULT ((0)) FOR Total
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioVenta DEFAULT ((0)) FOR PrecioVenta
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_TotalIngresos DEFAULT ((0)) FOR TotalIngresos
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Trazabilidad DEFAULT ((0)) FOR Trazabilidad
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_OtrasDeducciones DEFAULT ((0)) FOR OtrasDeducciones
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_Bolsa DEFAULT ((0)) FOR Bolsa
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioLunes DEFAULT ((0)) FOR PrecioLunes
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioMartes DEFAULT ((0)) FOR PrecioMartes
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioMiercoles DEFAULT ((0)) FOR PrecioMiercoles
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioJueves DEFAULT ((0)) FOR PrecioJueves
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioViernes DEFAULT ((0)) FOR PrecioViernes
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioSabado DEFAULT ((0)) FOR PrecioSabado
GO
ALTER TABLE dbo.Tmp_Detalle_Nomina ADD CONSTRAINT
	DF_Detalle_Nomina_PrecioDomingo DEFAULT ((0)) FOR PrecioDomingo
GO
IF EXISTS(SELECT * FROM dbo.Detalle_Nomina)
	 EXEC('INSERT INTO dbo.Tmp_Detalle_Nomina (NumNomina, CodProductor, TipoProductor, Roc1, Lunes, Roc2, Martes, Roc3, Miercoles, Roc4, Jueves, Roc5, Viernes, Roc6, Sabado, Roc7, Domingo, Total, PrecioVenta, TotalIngresos, IR, DeduccionPolicia, Anticipo, DeduccionTransporte, Pulperia, Inseminacion, ProductosVeterinarios, Trazabilidad, OtrasDeducciones, Nombres, Bolsa, PrecioLunes, PrecioMartes, PrecioMiercoles, PrecioJueves, PrecioViernes, PrecioSabado, PrecioDomingo)
		SELECT NumNomina, CodProductor, TipoProductor, Roc1, Lunes, Roc2, Martes, Roc3, Miercoles, Roc4, Jueves, Roc5, Viernes, Roc6, Sabado, Roc7, Domingo, Total, PrecioVenta, CONVERT(decimal(18, 2), TotalIngresos), CONVERT(decimal(18, 2), IR), CONVERT(decimal(18, 2), DeduccionPolicia), CONVERT(decimal(18, 2), Anticipo), CONVERT(decimal(18, 2), DeduccionTransporte), CONVERT(decimal(18, 2), Pulperia), CONVERT(decimal(18, 2), Inseminacion), CONVERT(decimal(18, 2), ProductosVeterinarios), CONVERT(decimal(18, 2), Trazabilidad), CONVERT(decimal(18, 2), OtrasDeducciones), Nombres, CONVERT(decimal(18, 2), Bolsa), PrecioLunes, PrecioMartes, PrecioMiercoles, PrecioJueves, PrecioViernes, PrecioSabado, PrecioDomingo FROM dbo.Detalle_Nomina WITH (HOLDLOCK TABLOCKX)')
GO
DROP TABLE dbo.Detalle_Nomina
GO
EXECUTE sp_rename N'dbo.Tmp_Detalle_Nomina', N'Detalle_Nomina', 'OBJECT' 
GO
ALTER TABLE dbo.Detalle_Nomina ADD CONSTRAINT
	PK_DetalleNomina PRIMARY KEY CLUSTERED 
	(
	NumNomina,
	CodProductor,
	TipoProductor
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.Detalle_Nomina ADD CONSTRAINT
	FK_DetalleNomina_Nomina FOREIGN KEY
	(
	NumNomina
	) REFERENCES dbo.Nomina
	(
	NumPlanilla
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
ALTER TABLE dbo.Detalle_Nomina ADD CONSTRAINT
	FK_DetalleNomina_Productor FOREIGN KEY
	(
	CodProductor,
	TipoProductor
	) REFERENCES dbo.Productor
	(
	CodProductor,
	TipoProductor
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
