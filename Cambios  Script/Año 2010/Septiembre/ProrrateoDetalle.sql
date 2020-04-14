/*
   Sábado, 05 de Septiembre de 2009 08:34:28 a.m.
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
CREATE TABLE dbo.DetalleProrrateo
	(
	NumeroProrrateo float(53) NOT NULL,
	CodCuenta nvarchar(50) NOT NULL,
	TipoProrrateo nvarchar(50) NULL,
	Descripcion nvarchar(50) NULL,
	MontoBase float(53) NULL,
	Porciento float(53) NULL,
	Importe float(53) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.DetalleProrrateo ADD CONSTRAINT
	PK_DetalleProrrateo PRIMARY KEY CLUSTERED 
	(
	NumeroProrrateo,
	CodCuenta
	) ON [PRIMARY]

GO
COMMIT
