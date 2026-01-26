SELECT Apellidos_Final.Apellido, Apellidos_Final.Idioma, Apellidos_Final.Tipo, Apellidos_Final.[Raiz etimológica]
FROM Apellidos_Final LEFT JOIN Apellidos_es ON Apellidos_Final.[Apellido] = Apellidos_es.[Apellido]
WHERE (((Apellidos_es.Apellido) Is Null));

