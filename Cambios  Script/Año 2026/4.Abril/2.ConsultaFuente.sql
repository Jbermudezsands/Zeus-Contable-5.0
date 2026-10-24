SELECT  
    Valor AS Fuente
FROM (
    -- TipoTransacciones (concatenado)
    SELECT 
        CodTipo + ' - ' + ISNULL(Descripcion, '') AS Valor
    FROM TipoTransacciones

    UNION

    -- Transacciones (Fuente)
    SELECT 
        Fuente
    FROM Transacciones
    WHERE Fuente IS NOT NULL
      AND LTRIM(RTRIM(Fuente)) <> ''
) AS Resultado
ORDER BY Fuente;