SELECT DISTINCT [Nombres EN-US].nombre, [Nombres EN-US].idioma, [Nombres EN-US].genero, [Nombres EN-US].notas
FROM [Nombres EN-US] LEFT JOIN tbmNombres ON ([Nombres EN-US].genero = tbmNombres.Genero) AND ([Nombres EN-US].idioma = tbmNombres.Idioma) AND ([Nombres EN-US].[nombre] = tbmNombres.[Nombre])
WHERE (((tbmNombres.Nombre) Is Null));

