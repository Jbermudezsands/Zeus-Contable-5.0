USE [SistemaContableNorteakMadera]
GO

/****** Object:  Table [dbo].[CatCombus]    Script Date: 10/31/2012 17:29:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CatCombus](
	[IdCombus] [int] IDENTITY(1,1) NOT NULL,
	[DescripCombus] [nvarchar](100) NULL,
 CONSTRAINT [PK_CatCombus] PRIMARY KEY CLUSTERED 
(
	[IdCombus] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


