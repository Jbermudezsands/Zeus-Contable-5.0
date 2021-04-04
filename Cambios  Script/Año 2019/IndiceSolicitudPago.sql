
GO

/****** Object:  Table [dbo].[IndiceTransaccion]    Script Date: 14/02/2019 08:36:36 p.m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[IndiceSolicitudPago](
	[FechaTransaccion] [smalldatetime] NOT NULL,
	[NumeroMovimiento] [int] NOT NULL,
	[DescripcionMovimiento] [nvarchar](255) NULL,
	[Nperiodo] [int] NULL,
	[Fuente] [nvarchar](255) NULL,
	[TipoMoneda] [nvarchar](50) NULL,
	[ImprimeCheque] [bit] NULL CONSTRAINT [DF_IndiceSolicitudPago_ImprimeCheque]  DEFAULT ((1)),
 CONSTRAINT [PK_IndiceSolicitudPago] PRIMARY KEY CLUSTERED 
(
	[FechaTransaccion] ASC,
	[NumeroMovimiento] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


