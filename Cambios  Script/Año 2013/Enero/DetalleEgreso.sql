/*
   sábado, 19 de enero de 201307:50:09 a.m.
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
CREATE TABLE dbo.DetalleEgreso
	(
	NumeroEgreso float(53) NOT NULL,
	CodCuentas nvarchar(50) NOT NULL,
	NombreCuenta nvarchar(MAX) NULL,
	Concepto nvarchar(MAX) NULL,
	ReciboNumero nvarchar(50) NULL,
	FacturaNumero nvarchar(50) NULL,
	FechaGasto smalldatetime NULL,
	MontoGasto numeric(18, 2) NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE dbo.DetalleEgreso ADD CONSTRAINT
	PK_DetalleEgreso PRIMARY KEY CLUSTERED 
	(
	NumeroEgreso,
	CodCuentas
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
