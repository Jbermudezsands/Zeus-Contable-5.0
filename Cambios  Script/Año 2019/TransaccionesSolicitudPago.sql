
GO

/****** Object:  Table [dbo].[Transacciones]    Script Date: 14/02/2019 08:39:34 p.m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TransaccionesSolicitudPago](
	[CodCuentas] [nvarchar](50) NOT NULL,
	[FechaTransaccion] [smalldatetime] NOT NULL,
	[NPeriodo] [int] NOT NULL,
	[NTransaccion] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[NumeroMovimiento] [int] NULL,
	[NombreCuenta] [nvarchar](255) NULL,
	[VoucherNo] [nvarchar](50) NULL,
	[DescripcionMovimiento] [nvarchar](255) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_DescripcionMovimiento]  DEFAULT (N'*'),
	[Clave] [nvarchar](10) NULL,
	[TCambio] [float] NULL,
	[Debito] [decimal](18, 2) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_Debito]  DEFAULT ((0)),
	[Credito] [decimal](18, 2) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_Credito]  DEFAULT ((0)),
	[FacturaNo] [nvarchar](50) NULL,
	[ChequeNo] [nvarchar](50) NULL,
	[Fuente] [nvarchar](50) NULL,
	[FechaTasas] [smalldatetime] NULL,
	[Conciliada] [int] NULL CONSTRAINT [DF_TransaccionesSolicitudPago_Conciliada]  DEFAULT ((0)),
	[ConciliacionProcesada] [int] NULL CONSTRAINT [DF_TransaccionesSolicitudPago_ConciliacionProcesada]  DEFAULT ((0)),
	[FechaDescuento] [smalldatetime] NULL CONSTRAINT [DF_TransaccionesSolicitudPago_FechaDescuento]  DEFAULT (((1)/(1))/(1900)),
	[DescuentoDisponible] [money] NULL CONSTRAINT [DF_TransaccionesSolicitudPago_DescuentoDisponible]  DEFAULT ((0)),
	[FechaVence] [smalldatetime] NULL CONSTRAINT [DF_TransaccionesSolicitudPago_FechaVence]  DEFAULT (((1)/(1))/(1900)),
	[Beneficiario] [nvarchar](255) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_Beneficiario]  DEFAULT (N'-'),
	[CodCuentaProveedor] [nvarchar](50) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_CodCuentaProveedor]  DEFAULT ((0)),
	[TipoFactura] [nvarchar](50) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_TipoFactura]  DEFAULT (N'-'),
	[DebitoD] [decimal](18, 2) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_DebitoD]  DEFAULT ((0)),
	[CreditoD] [decimal](18, 2) NULL CONSTRAINT [DF_TransaccionesSolicitudPago_CreditoD]  DEFAULT ((0)),
 CONSTRAINT [PK_TransaccionesSolicitudPago] PRIMARY KEY CLUSTERED 
(
	[CodCuentas] ASC,
	[FechaTransaccion] ASC,
	[NPeriodo] ASC,
	[NTransaccion] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


