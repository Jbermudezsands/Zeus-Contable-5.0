USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[ControlOrdenTrabajo]    Script Date: 10/31/2012 17:31:05 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ControlOrdenTrabajo](
	[idreg] [int] IDENTITY(1,1) NOT NULL,
	[BienOrden] [nvarchar](120) NULL,
	[Fcreado] [smalldatetime] NULL,
	[Reportadopor] [nvarchar](50) NULL,
	[frequeireOrden] [smalldatetime] NULL,
	[proveeresponsable] [nvarchar](50) NULL,
	[Descripcion] [nvarchar](200) NULL,
	[Estado] [nvarchar](1) NULL,
	[referencia] [nvarchar](30) NULL,
	[Nota] [nvarchar](200) NULL,
	[IdActivo] [int] NULL,
 CONSTRAINT [PK_ControlOrdenTrabajo] PRIMARY KEY CLUSTERED 
(
	[idreg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


