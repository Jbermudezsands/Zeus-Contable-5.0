/*
   viernes, 19 de julio de 201309:55:10 p.m.
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
CREATE TABLE dbo.AltadeBienes
	(
	IdReg int NOT NULL IDENTITY (1, 1),
	IdReferencia nvarchar(50) NULL,
	FechaGraba smalldatetime NULL,
	IdOfiDestino int NULL,
	DescriOficina nvarchar(50) NULL,
	Observaciones nvarchar(MAX) NULL,
	IdUserRecibe int NULL,
	NombreRecibe nvarchar(50) NULL,
	IdUserEntrega int NULL,
	NombreEntrega nvarchar(50) NULL,
	IdActivoAlta nvarchar(50) NULL,
	IdOfiAlta int NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE dbo.AltadeBienes ADD CONSTRAINT
	PK_AltadeBienes PRIMARY KEY CLUSTERED 
	(
	IdReg
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
