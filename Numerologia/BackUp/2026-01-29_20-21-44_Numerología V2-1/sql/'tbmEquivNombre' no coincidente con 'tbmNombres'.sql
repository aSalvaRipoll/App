SELECT Nombre, Idioma
FROM tbmNombres
ORDER BY Nombre
UNION SELECT DISTINCT tbmEquivNombre.NombreOriginal AS Nombre, tbmEquivNombre.IdiomaOriginal
FROM tbmEquivNombre;

