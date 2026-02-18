SELECT DISTINCT Nombre, Idioma
FROM tbmNombres
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_01
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_02
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_03
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_04
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_05
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_06
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_07
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_08
UNION ALL
SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_09
UNION ALL SELECT DISTINCT NombreOriginal, IdiomaOriginal
FROM EquivalenciasNombres_10;

