/*
   viernes, 19 de julio de 201310:37:45 p.m.
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
CREATE TABLE dbo.CatalogoActivoFijo
	(
	idReg int NOT NULL IDENTITY (1, 1),
	Unidad int NULL,
	marca nvarchar(250) NULL,
	modelo nvarchar(250) NULL,
	año nvarchar(250) NULL,
	color nvarchar(250) NULL,
	DescripcionAF nvarchar(250) NULL,
	Serie nvarchar(50) NULL,
	tipovehiculo int NULL,
	descriptipoveh nvarchar(150) NULL,
	tipocombus int NULL,
	descriptipocombus nvarchar(150) NULL,
	grupo int NULL,
	descrigrupo nvarchar(150) NULL,
	codconductor int NULL,
	nombreconduc nvarchar(150) NULL,
	placa nvarchar(50) NULL,
	frenovacion smalldatetime NULL,
	nota nvarchar(350) NULL,
	isvehipropio int NULL,
	fadquicisionvh smalldatetime NULL,
	kilomcompravh numeric(18, 2) NULL,
	compradooalqui nvarchar(50) NULL,
	costovh numeric(18, 2) NULL,
	ivavh numeric(18, 2) NULL,
	garantiacaduvh smalldatetime NULL,
	notacompravh nvarchar(350) NULL,
	Aseguradorvh nvarchar(150) NULL,
	compasegvh nvarchar(150) NULL,
	referencia nvarchar(150) NULL,
	finiasevh smalldatetime NULL,
	ffinasevh smalldatetime NULL,
	notaasevh nvarchar(MAX) NULL,
	perrefvh nvarchar(50) NULL,
	notapervh nvarchar(250) NULL,
	finiper smalldatetime NULL,
	ffinper smalldatetime NULL,
	alarmaseguro int NULL,
	alarmapermiso int NULL,
	cntacontable nvarchar(50) NULL,
	refegeneral nvarchar(50) NULL,
	factura nvarchar(50) NULL,
	fcompragen smalldatetime NULL,
	costogen numeric(18, 2) NULL,
	ivagen numeric(18, 2) NULL,
	FechaBaja nvarchar(50) NULL,
	DatoAlta bit NULL,
	fechaalta smalldatetime NULL,
	dadobaja bit NULL,
	idofialta int NULL,
	trasladado bit NULL,
	fechatraslado smalldatetime NULL,
	CuentaGastos nvarchar(50) NULL,
	CuentaDepreciacion nvarchar(50) NULL,
	isvh bit NULL,
	dirfoto nvarchar(MAX) NULL,
	dirfoto1 nvarchar(MAX) NULL,
	dirfoto2 nvarchar(MAX) NULL,
	IdActivoAlta int NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	DF_CatalogoActivoFijo_DatoAlta DEFAULT 0 FOR DatoAlta
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	DF_CatalogoActivoFijo_dadobaja DEFAULT 0 FOR dadobaja
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	DF_CatalogoActivoFijo_trasladado DEFAULT 0 FOR trasladado
GO
ALTER TABLE dbo.CatalogoActivoFijo ADD CONSTRAINT
	PK_CatalogoActivoFijo PRIMARY KEY CLUSTERED 
	(
	idReg
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
