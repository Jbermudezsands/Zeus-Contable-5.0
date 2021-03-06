/*
   viernes, 19 de julio de 201310:01:47 p.m.
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
CREATE TABLE dbo.BajadeBienes
	(
	IdReg int NOT NULL IDENTITY (1, 1),
	IdReferencia nvarchar(50) NULL,
	FechaGraba smalldatetime NULL,
	IdOfiOrigen int NULL,
	DescriOficina nvarchar(50) NULL,
	Observaciones nvarchar(MAX) NULL,
	IdUserRecibe int NULL,
	NombreRecibe nvarchar(50) NULL,
	IdUserEntrega int NULL,
	NombreEntrega nvarchar(50) NULL,
	IdUserAutoriza int NULL,
	NombreAutoriza nvarchar(50) NULL,
	IdActivoBaja nvarchar(50) NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE dbo.BajadeBienes ADD CONSTRAINT
	PK_BajadeBienes PRIMARY KEY CLUSTERED 
	(
	IdReg
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
