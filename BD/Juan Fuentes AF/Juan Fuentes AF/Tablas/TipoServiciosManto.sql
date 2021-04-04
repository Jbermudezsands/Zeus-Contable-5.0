USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[TipoServiciosManto]    Script Date: 10/31/2012 17:31:32 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TipoServiciosManto](
	[idreg] [int] IDENTITY(1,1) NOT NULL,
	[repetircada] [int] NULL,
	[mostraralerta] [int] NULL,
	[DescripcionMantto] [nvarchar](150) NULL,
	[TipoRepeticion] [nvarchar](20) NULL,
	[TipoRepAlerta] [nvarchar](20) NULL,
 CONSTRAINT [PK_TipoServiciosManto] PRIMARY KEY CLUSTERED 
(
	[idreg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


