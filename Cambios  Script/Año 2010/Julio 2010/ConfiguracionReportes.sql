/*
   Jueves, 22 de Julio de 2010 04:15:45 p.m.
   Usuario: 
   Servidor: JUAN\SQL2000
   Base de datos: SistemaContableDemo
   Aplicación: MS SQLEM - Data Tools
*/

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
CREATE TABLE dbo.ConfiguracionReporte
	(
	Caja nvarchar(50) NULL,
	Banco nvarchar(50) NULL,
	CtasxCobrar nvarchar(50) NULL,
	Inventario nvarchar(50) NULL,
	Terreno nvarchar(50) NULL,
	Mobiliario nvarchar(50) NULL,
	EquipoRodante nvarchar(50) NULL,
	DepAcumulada nvarchar(50) NULL,
	Papeleria nvarchar(50) NULL,
	UtiliesOficinas nvarchar(50) NULL,
	PagosAnticipados nvarchar(50) NULL,
	OtrosActivos nvarchar(50) NULL,
	Proveedores nvarchar(50) NULL,
	ImpuestosxPagar nvarchar(50) NULL,
	DocumentosxPagar nvarchar(50) NULL,
	CobroAnticipados nvarchar(50) NULL,
	PasivosAcumulados nvarchar(50) NULL,
	PagosLP nvarchar(50) NULL,
	DocumentosLP nvarchar(50) NULL,
	DepreciacionAcum nvarchar(50) NULL,
	OtrosPasivos nvarchar(50) NULL,
	AccionesComunes nvarchar(50) NULL,
	UtilidadAcumulada nvarchar(50) NULL,
	OtrosCapitales nvarchar(50) NULL,
	AnexoCaja bit NULL,
	AnexoBanco bit NULL,
	AnexoCtasxCobrar bit NULL,
	AnexoInventario bit NULL,
	AnexoTerreno bit NULL,
	AnexoMobiliario bit NULL,
	AnexoEquipoRodante bit NULL,
	AnexoDepAcumulada bit NULL,
	AnexoPapeleria bit NULL,
	AnexoUtilesOficina bit NULL,
	AnexosPagosAnticipados bit NULL,
	AnexosOtrosActivos bit NULL,
	AnexosProveedores bit NULL,
	AnexosImpuestosxPagar bit NULL,
	AnexosDocumentosxPagar bit NULL,
	AnexosCobroAnticipados bit NULL,
	AnexosPasivosAcumulados bit NULL,
	AnexosPagosLP bit NULL,
	AnexosDocumentosLP bit NULL,
	AnexosDepreciacionAcum bit NULL,
	AnexosOtrosPasivos bit NULL,
	AnexosAccionesComunes bit NULL,
	AnexosUtilidadAcumulada bit NULL,
	AnexosOtrosCapitales bit NULL
	)  ON [PRIMARY]
GO
COMMIT
