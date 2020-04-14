USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[CatalogoActivoFijo]    Script Date: 10/31/2012 17:28:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CatalogoActivoFijo](
	[idReg] [int] IDENTITY(1,1) NOT NULL,
	[Unidad] [int] NULL,
	[marca] [nvarchar](250) NULL,
	[modelo] [nvarchar](250) NULL,
	[año] [int] NULL,
	[color] [nvarchar](25) NULL,
	[DescripcionAF] [nvarchar](250) NULL,
	[Serie] [nvarchar](45) NULL,
	[tipovehiculo] [int] NULL,
	[descriptipoveh] [nvarchar](90) NULL,
	[tipocombus] [int] NULL,
	[descriptipocombus] [nvarchar](90) NULL,
	[grupo] [int] NULL,
	[descrigrupo] [nvarchar](90) NULL,
	[codconductor] [int] NULL,
	[nombreconduc] [nvarchar](90) NULL,
	[placa] [nvarchar](15) NULL,
	[frenovacion] [smalldatetime] NULL,
	[nota] [nvarchar](300) NULL,
	[isvehipropio] [int] NULL,
	[fadquicisionvh] [smalldatetime] NULL,
	[kilomcompravh] [numeric](18, 2) NULL,
	[compradooalqui] [nvarchar](50) NULL,
	[costovh] [money] NULL,
	[ivavh] [money] NULL,
	[garantiacaduvh] [smalldatetime] NULL,
	[notacompravh] [nvarchar](300) NULL,
	[Aseguradorvh] [nvarchar](100) NULL,
	[compasegvh] [nvarchar](100) NULL,
	[referencia] [nvarchar](50) NULL,
	[finiasevh] [smalldatetime] NULL,
	[ffinasevh] [smalldatetime] NULL,
	[notaasevh] [nvarchar](300) NULL,
	[perrefvh] [nvarchar](50) NULL,
	[notapervh] [nvarchar](200) NULL,
	[finiper] [smalldatetime] NULL,
	[ffinper] [smalldatetime] NULL,
	[alarmaseguro] [int] NULL,
	[alarmapermiso] [int] NULL,
	[cntacontable] [nvarchar](20) NULL,
	[refegeneral] [nvarchar](50) NULL,
	[factura] [nvarchar](20) NULL,
	[fcompragen] [smalldatetime] NULL,
	[costogen] [money] NULL,
	[ivagen] [money] NULL,
	[FechaBaja] [smalldatetime] NULL,
	[DatoAlta] [bit] NULL,
	[fechaalta] [smalldatetime] NULL,
	[dadobaja] [bit] NULL,
	[idofialta] [int] NULL,
	[trasladado] [bit] NULL,
	[fechatraslado] [smalldatetime] NULL,
	[CuentaGastos] [nvarchar](50) NULL,
	[CuentaDepreciacion] [nvarchar](50) NULL,
	[isvh] [bit] NULL,
	[dirfoto] [nvarchar](200) NULL,
	[dirfoto1] [nvarchar](200) NULL,
	[dirfoto2] [nvarchar](200) NULL,
	[IdActivoAlta] [int] NULL,
 CONSTRAINT [PK_CatalogoActivoFijo] PRIMARY KEY CLUSTERED 
(
	[idReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


