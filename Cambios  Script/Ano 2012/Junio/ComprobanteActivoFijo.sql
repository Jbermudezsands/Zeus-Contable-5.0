/*
   domingo, 01 de julio de 201204:40:28 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaContableSystems
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
CREATE TABLE dbo.ComprobanteActivoFijo
	(
	CodActivo nvarchar(50) NOT NULL,
	FechaCalculo smalldatetime NOT NULL,
	Cuenta nvarchar(50) NULL,
	Descripcion nvarchar(MAX) NULL,
	Debe float(53) NULL,
	Haber float(53) NULL,
	Contabilizado bit NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE dbo.ComprobanteActivoFijo ADD CONSTRAINT
	DF_ComprobanteActivoFijo_Contabilizado DEFAULT 0 FOR Contabilizado
GO
ALTER TABLE dbo.ComprobanteActivoFijo ADD CONSTRAINT
	PK_ComprobanteActivoFijo PRIMARY KEY CLUSTERED 
	(
	CodActivo,
	FechaCalculo
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
