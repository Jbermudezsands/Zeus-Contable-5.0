/*
   sábado, 19 de agosto de 202302:57:50 p.m.
   User: 
   Server: JUANBERMUDEZ\SQL2019
   Database: SistemaContableEmtrides
   Application: 
*/

/* To prevent any potential data loss issues, you should review this script in detail before running it outside the context of the database designer.*/
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
CREATE TABLE dbo.Registros_PBI
	(
	Numero_Registro float(53) NULL,
	Fecha_Registro smalldatetime NULL,
	Dia float(53) NULL,
	Mes float(53) NULL,
	Año float(53) NULL,
	Moneda nvarchar(50) NULL,
	Tipo_Cambio float(53) NULL,
	Tipo_Movimiento nvarchar(50) NULL,
	DEBE float(53) NULL,
	HABER float(53) NULL,
	SALDO float(53) NULL,
	Cod_Cuenta nvarchar(50) NULL,
	Descripcion_Cuenta nvarchar(250) NULL,
	Tipo_Cuenta nvarchar(50) NULL,
	Fuente nvarchar(50) NULL,
	Causal_Registro nvarchar(250) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Registros_PBI SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
