SELECT tbmEquivNombre_3.NombreOriginal, EtimologiaNombres.Origen, EtimologiaNombres.Etimologia, EtimologiaNombres.Campo5
FROM tbmEquivNombre_3 LEFT JOIN EtimologiaNombres ON (tbmEquivNombre_3.Genero = EtimologiaNombres.Genero) AND (tbmEquivNombre_3.NombreOriginal = EtimologiaNombres.Nombre)
WHERE (((tbmEquivNombre_3.NombreOriginal) Like "alc*"));

