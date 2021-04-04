

/****** Object:  Table [dbo].[Grupos]    Script Date: 22/08/2020 08:45:39 a.m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[EstructuraPresupuesto](
	[KeyGrupo] [nvarchar](255) NOT NULL,
	[CodGrupo] [nvarchar](55) NULL,
	[KeyGrupoSuperior] [nvarchar](255) NULL,
	[Child] [nvarchar](50) NULL,
	[DescripcionGrupo] [nvarchar](255) NULL,
	[Imagen1] [int] NULL,
	[Imagen2] [int] NULL,
 CONSTRAINT [PK_EstructuraPresupuesto] PRIMARY KEY CLUSTERED 
(
	[KeyGrupo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


