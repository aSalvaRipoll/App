SELECT Consulta12.NombreOriginal, Consulta12.IdiomaOriginal, Consulta12.genero
FROM Consulta12 LEFT JOIN [Nombres EN-US] ON (Consulta12.genero = [Nombres EN-US].genero) AND (Consulta12.IdiomaOriginal = [Nombres EN-US].idioma) AND (Consulta12.[NombreOriginal] = [Nombres EN-US].[nombre])
WHERE ((([Nombres EN-US].nombre) Is Null));

