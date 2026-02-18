SELECT tbmEquivNombre_3.NombreOriginal, tbmEquivNombre_3.IdiomaOriginal, tbmEquivNombre_3.NombreEquivalente, tbmEquivNombre_3.IdiomaEquivalente, tbmEquivNombre_3.Genero, tbmEquivNombre_3.Activo
FROM tbmEquivNombre_3 INNER JOIN tbmEtimologiaNombres ON (tbmEquivNombre_3.Genero = tbmEtimologiaNombres.Genero) AND (tbmEquivNombre_3.NombreOriginal = tbmEtimologiaNombres.Nombre)
UNION SELECT tbmEquivNombre_3.NombreEquivalente, tbmEquivNombre_3.IdiomaEquivalente, tbmEquivNombre_3.NombreOriginal, tbmEquivNombre_3.IdiomaOriginal, tbmEquivNombre_3.Genero, tbmEquivNombre_3.Activo
FROM tbmEquivNombre_3 INNER JOIN tbmEtimologiaNombres ON (tbmEquivNombre_3.Genero = tbmEtimologiaNombres.Genero) AND (tbmEquivNombre_3.NombreEquivalente = tbmEtimologiaNombres.Nombre);

