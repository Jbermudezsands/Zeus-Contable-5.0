SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ResponsablesAreas](
	[IdReg] [int] IDENTITY(1,1) NOT NULL,
	[NombreResponsable] [nvarchar](70) COLLATE Modern_Spanish_CI_AS NULL,
	[Area] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[Telefono] [nvarchar](10) COLLATE Modern_Spanish_CI_AS NULL,
	[email] [nvarchar](40) COLLATE Modern_Spanish_CI_AS NULL,
	[Cargo] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[FechaReg] [smalldatetime] NULL,
	[IdAreaTrabajo] [int] NULL,
	[Cedula] [nvarchar](18) COLLATE Modern_Spanish_CI_AS NULL,
 CONSTRAINT [PK_ResponsablesAreas] PRIMARY KEY CLUSTERED 
(
	[IdReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
