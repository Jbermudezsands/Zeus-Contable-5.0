/*
   jueves, 23 de mayo de 201301:00:53 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaContableALFA
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar esta secuencia de comandos detalladamente antes de ejecutarla fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.ConfiguracionReporte ADD
	IngresosVentas nvarchar(50) NULL,
	ServiciosVentas nvarchar(50) NULL,
	ComisionVentas nvarchar(50) NULL,
	RebajayDevolucionesVentas nvarchar(50) NULL,
	CostodeVentas nvarchar(50) NULL,
	CostodeProduccion nvarchar(50) NULL,
	CostosGeneralesdeProduccion nvarchar(50) NULL,
	SueldosyComisiones nvarchar(50) NULL,
	Propaganda nvarchar(50) NULL,
	Sueldos nvarchar(50) NULL,
	EnergiaElectrica nvarchar(50) NULL,
	ComisionesGanadas nvarchar(50) NULL,
	ComisionesPagadas nvarchar(50) NULL,
	OtrosIngresosyGastos nvarchar(50) NULL,
	AnexosIngresosVentas bit NULL,
	AnexosServiciosVentas bit NULL,
	AnexosComisionVentas bit NULL,
	AnexosRebajasyDevolucionesVentas bit NULL,
	AnexosCostosdeVentas bit NULL,
	AnexosCostosdeProduccionAnexosCostosGeneralesdeProduccion bit NULL,
	AnexosSueldosyComisiones bit NULL,
	AnexosPropaganda bit NULL,
	AnexosSueldos bit NULL,
	AnexosEnergiaElectrica bit NULL,
	AnexosComisionesGanadas bit NULL,
	AnexosComisionesPagadas bit NULL,
	AnexosOtrosIngresosyGastos bit NULL
GO
COMMIT
