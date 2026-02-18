SELECT tbmEquivNombre_2.NombreOriginal, tbmEquivNombre_2.IdiomaOriginal, tbmEquivNombre_2.NombreEquivalente, tbmEquivNombre_2.IdiomaEquivalente, EtimologiaNombres.Origen, EtimologiaNombres.Etimologia, EtimologiaNombres.Genero
FROM tbmEquivNombre_2 INNER JOIN EtimologiaNombres ON tbmEquivNombre_2.NombreOriginal = EtimologiaNombres.Nombre
WHERE (((EtimologiaNombres.Genero)="M")) OR (((EtimologiaNombres.Genero)="F"));

