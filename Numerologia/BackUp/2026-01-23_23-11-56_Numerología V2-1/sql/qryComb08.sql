SELECT NombreOriginal, IdiomaOriginal, NombreEquivalente, IdiomaEquivalente, Tipo, Notas
FROM EquivalenciasNombres_08
UNION ALL SELECT NombreEquivalente, IdiomaEquivalente, NombreOriginal, IdiomaOriginal, Tipo, Notas
FROM EquivalenciasNombres_08;

