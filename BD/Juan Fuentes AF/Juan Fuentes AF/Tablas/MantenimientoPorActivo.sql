USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[MantenimientoPorActivo]    Script Date: 10/31/2012 17:31:53 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MantenimientoPorActivo](
	[idreg] [int] IDENTITY(1,1) NOT NULL,
	[IdActivo] [int] NULL,
	[IdServici] [int] NULL,
	[repetircada] [int] NULL,
	[mostraralerta] [int] NULL,
	[proximomanto] [smalldatetime] NULL,
	[Descripcion] [nvarchar](50) NULL,
	[tiporepeticion] [nvarchar](20) NULL,
	[TipoAlerta] [nvarchar](20) NULL,
	[ultimoservicio] [smalldatetime] NULL,
 CONSTRAINT [PK_MantenimientoPorActivo] PRIMARY KEY CLUSTERED 
(
	[idreg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


