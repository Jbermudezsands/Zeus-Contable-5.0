USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[ResponsablesAreas]    Script Date: 10/31/2012 17:28:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ResponsablesAreas](
	[IdReg] [int] IDENTITY(1,1) NOT NULL,
	[NombreResponsable] [nvarchar](70) NULL,
	[Area] [nvarchar](50) NULL,
	[Telefono] [nvarchar](10) NULL,
	[email] [nvarchar](40) NULL,
	[Cargo] [nvarchar](50) NULL,
	[FechaReg] [smalldatetime] NULL,
	[IdAreaTrabajo] [int] NULL,
	[Cedula] [nvarchar](18) NULL,
 CONSTRAINT [PK_ResponsablesAreas] PRIMARY KEY CLUSTERED 
(
	[IdReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


