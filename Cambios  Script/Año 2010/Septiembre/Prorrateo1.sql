/*
   Sábado, 05 de Septiembre de 2009 08:13:26 a.m.
   Usuario: 
   Servidor: JUAN\SQL2000
   Base de datos: SistemaContableMarAzul
   Aplicación: MS SQLEM - Data Tools
*/

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
CREATE TABLE dbo.Prorrateo
	(
	NumeroProrrateo float(53) NOT NULL,
	FechaMovimiento datetime NULL,
	PeriodoIni datetime NULL,
	PeriodoFin datetime NULL,
        Periodo1 float(53) NULL,
        Periodo2 float(53) NULL,
	Año float(53) NULL,
	NumeroTabla float(53) NULL,
	DescripcionMovimiento nvarchar(50) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Prorrateo ADD CONSTRAINT
	PK_Prorrateo PRIMARY KEY CLUSTERED 
	(
	NumeroProrrateo
	) ON [PRIMARY]

GO
COMMIT
