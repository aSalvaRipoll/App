SELECT tbmNombresEs.Nombre, tbmNombresEs.Idioma, tbmNombresEs.Genero
FROM tbmNombresEs LEFT JOIN tbmEqNombres ON tbmNombresEs.[Nombre] = tbmEqNombres.[NombreOriginal]
WHERE (((tbmNombresEs.Idioma)="es") AND ((tbmEqNombres.NombreOriginal) Is Null));

