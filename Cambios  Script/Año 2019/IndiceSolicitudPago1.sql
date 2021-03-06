/*
   jueves, 14 de febrero de 201909:12:27 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: SistemContablePanam
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
ALTER TABLE dbo.IndiceSolicitudPago ADD
	SubTotal float(53) NULL,
	MontoIva float(53) NULL,
	MontoRetenciones float(53) NULL,
	MontoSolicitud float(53) NULL,
	Anticipo bit NULL,
	Retencion1 bit NULL,
	Retencion2 bit NULL,
	Retencion3 bit NULL,
	Retencion4 bit NULL,
	Retencion5 bit NULL,
	Retencion6 bit NULL,
	Iva bit NULL,
	Concepto nvarchar(250) NULL
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_SubTotal DEFAULT 0 FOR SubTotal
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_MontoIva DEFAULT 0 FOR MontoIva
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_MontoRetenciones DEFAULT 0 FOR MontoRetenciones
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_MontoSolicitud DEFAULT 0 FOR MontoSolicitud
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Anticipo DEFAULT 0 FOR Anticipo
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Retencion1 DEFAULT 0 FOR Retencion1
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Retencion2 DEFAULT 0 FOR Retencion2
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Retencion3 DEFAULT 0 FOR Retencion3
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Retencion4 DEFAULT 0 FOR Retencion4
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Retencion5 DEFAULT 0 FOR Retencion5
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Retencion6 DEFAULT 0 FOR Retencion6
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Iva DEFAULT 0 FOR Iva
GO
ALTER TABLE dbo.IndiceSolicitudPago SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
