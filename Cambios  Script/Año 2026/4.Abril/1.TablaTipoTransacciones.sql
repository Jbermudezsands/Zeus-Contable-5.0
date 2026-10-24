
/****** Object:  Table [dbo].[TipoTransacciones]    Script Date: 08/04/2026 11:12:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TipoTransacciones](
	[CodTipo] [nvarchar](50) NOT NULL,
	[Descripcion] [nvarchar](250) NULL,
 CONSTRAINT [PK_TipoTransacciones] PRIMARY KEY CLUSTERED 
(
	[CodTipo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO


