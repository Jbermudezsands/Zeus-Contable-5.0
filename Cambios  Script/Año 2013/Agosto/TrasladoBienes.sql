SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TrasladoBienes](
	[IdReg] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[IdReferencia] [nvarchar](20) COLLATE Modern_Spanish_CI_AS NULL,
	[FechaGraba] [smalldatetime] NULL,
	[IdOfiOrigen] [int] NULL,
	[DescriOficina] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[IdOfiDestino] [int] NULL,
	[DescriOficinaDest] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[Observaciones] [nvarchar](200) COLLATE Modern_Spanish_CI_AS NULL,
	[IdUserRecibe] [int] NULL,
	[NombreRecibe] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[IdUserEntrega] [int] NULL,
	[NombreEntrega] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[IdUserAutoriza] [int] NULL,
	[NombreAutoriza] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[IdActivoTraslada] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
 CONSTRAINT [PK_TrasladoBienes] PRIMARY KEY CLUSTERED 
(
	[IdReg] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
