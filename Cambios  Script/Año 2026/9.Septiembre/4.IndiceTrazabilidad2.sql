CREATE INDEX IX_TrazabilidadContable_Documento
ON dbo.TrazabilidadContable
(
    ModuloOrigen,
    TipoDocumento,
    IdDocumento
);