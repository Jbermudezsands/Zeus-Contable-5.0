CREATE TABLE [DetalleMemoriza] (
	[CodCuentas] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[FechaTransaccion] [smalldatetime] NOT NULL ,
	[NPeriodo] [int] NOT NULL ,
	[NTransaccion] [int] IDENTITY (1, 1) NOT NULL ,
	[IdMemoria] int NOT NULL ,
	[NombreCuenta] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,
	[VoucherNo] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DescripcionMovimiento] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL CONSTRAINT [DF_DetalleMemoriza_DescripcionMovimiento] DEFAULT (N'*'),
	[Clave] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[TCambio] [float] NULL ,
	[Debito] [money] NULL CONSTRAINT [DF_DetalleMemoriza_Debito] DEFAULT (0),
	[Credito] [money] NULL CONSTRAINT [DF_DetalleMemoriza_Credito] DEFAULT (0),
	[FacturaNo] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ChequeNo] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fuente] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaTasas] [smalldatetime] NULL ,
	[Conciliada] [int] NULL CONSTRAINT [DF_DetalleMemoriza_Conciliada] DEFAULT (0),
	[ConciliacionProcesada] [int] NULL CONSTRAINT [DF_DetalleMemoriza_ConciliacionProcesada] DEFAULT (0),
	[FechaDescuento] [smalldatetime] NULL CONSTRAINT [DF_DetalleMemoriza_FechaDescuento] DEFAULT (1 / 1 / 1900),
	[DescuentoDisponible] [money] NULL CONSTRAINT [DF_DetalleMemoriza_DescuentoDisponible] DEFAULT (0),
	[FechaVence] [smalldatetime] NULL CONSTRAINT [DF_DetalleMemoriza_FechaVence] DEFAULT (1 / 1 / 1900),
	[Beneficiario] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL CONSTRAINT [DF_DetalleMemoriza_Beneficiario] DEFAULT (N'-'),
	[CodCuentaProveedor] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL CONSTRAINT [DF_DetalleMemoriza_CodCuentaProveedor] DEFAULT (0),
	[TipoFactura] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL CONSTRAINT [DF_DetalleMemoriza_TipoFactura] DEFAULT (N'-'),
	CONSTRAINT [PK_DetalleMemoriza] PRIMARY KEY  CLUSTERED 
	(
		[CodCuentas],
		[FechaTransaccion],
		[NPeriodo],
		[NTransaccion]
	)  ON [PRIMARY] 
) ON [PRIMARY]
GO


