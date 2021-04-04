USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[TrasladoBienes]    Script Date: 10/31/2012 17:27:19 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TrasladoBienes](
	[IdReg] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[IdReferencia] [nvarchar](20) NULL,
	[FechaGraba] [smalldatetime] NULL,
	[IdOfiOrigen] [int] NULL,
	[DescriOficina] [nvarchar](50) NULL,
	[IdOfiDestino] [int] NULL,
	[DescriOficinaDest] [nvarchar](50) NULL,
	[Observaciones] [nvarchar](200) NULL,
	[IdUserRecibe] [int] NULL,
	[NombreRecibe] [nvarchar](50) NULL,
	[IdUserEntrega] [int] NULL,
	[NombreEntrega] [nvarchar](50) NULL,
	[IdUserAutoriza] [int] NULL,
	[NombreAutoriza] [nvarchar](50) NULL,
	[IdActivoTraslada] [nvarchar](50) NULL,
 CONSTRAINT [PK_TrasladoBienes] PRIMARY KEY CLUSTERED 
(
	[IdReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


