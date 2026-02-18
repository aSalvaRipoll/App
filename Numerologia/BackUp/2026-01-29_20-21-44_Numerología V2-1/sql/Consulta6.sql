SELECT tbmNombres.Nombre, tbmNombres.Idioma, EtimologiaNombres.Origen, EtimologiaNombres.Etimologia, tbmNombres.Notas, tbmNombres.Activo
FROM EtimologiaNombres INNER JOIN tbmNombres ON (EtimologiaNombres.Genero = tbmNombres.Genero) AND (EtimologiaNombres.Nombre = tbmNombres.Nombre);

