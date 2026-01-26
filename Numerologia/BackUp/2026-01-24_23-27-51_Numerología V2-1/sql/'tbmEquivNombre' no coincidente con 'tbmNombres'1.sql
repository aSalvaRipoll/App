SELECT DISTINCT tbmEquivNombre.NombreEquivalente, tbmEquivNombre.IdiomaEquivalente
FROM tbmEquivNombre LEFT JOIN tbmNombres ON (tbmEquivNombre.IdiomaEquivalente = tbmNombres.Idioma) AND (tbmEquivNombre.[NombreEquivalente] = tbmNombres.[Nombre])
WHERE (((tbmNombres.Nombre) Is Null)) OR (((tbmNombres.Idioma) Is Null));

