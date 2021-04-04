USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[CataVH]    Script Date: 10/31/2012 17:28:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CataVH](
	[idvh] [int] IDENTITY(1,1) NOT NULL,
	[descricpcion] [nvarchar](100) NULL,
 CONSTRAINT [PK_CataVH] PRIMARY KEY CLUSTERED 
(
	[idvh] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


