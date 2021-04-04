USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[AltadeBienes]    Script Date: 10/31/2012 17:26:05 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[AltadeBienes](
	[IdReg] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[IdReferencia] [nvarchar](20) NULL,
	[FechaGraba] [smalldatetime] NULL,
	[IdOfiDestino] [int] NULL,
	[DescriOficina] [nvarchar](50) NULL,
	[Observaciones] [nvarchar](200) NULL,
	[IdUserRecibe] [int] NULL,
	[NombreRecibe] [nvarchar](50) NULL,
	[IdUserEntrega] [int] NULL,
	[NombreEntrega] [nvarchar](50) NULL,
	[IdActivoAlta] [nvarchar](50) NULL,
	[IdOfiAlta] [int] NULL,
 CONSTRAINT [PK_AltadeBienes] PRIMARY KEY CLUSTERED 
(
	[IdReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


