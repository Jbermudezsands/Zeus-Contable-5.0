CREATE INDEX IX_TrazabilidadContable_Movimiento
ON dbo.TrazabilidadContable
(
    FechaTransaccion,
    NPeriodo,
    NumeroMovimiento
);