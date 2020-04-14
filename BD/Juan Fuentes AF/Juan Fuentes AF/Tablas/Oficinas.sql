USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[Oficinas]    Script Date: 10/31/2012 17:27:50 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Oficinas](
	[IdReg] [int] IDENTITY(1,1) NOT NULL,
	[Descripcion] [nvarchar](50) NULL,
	[FechaReg] [smalldatetime] NULL,
 CONSTRAINT [PK_Oficinas] PRIMARY KEY CLUSTERED 
(
	[IdReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


