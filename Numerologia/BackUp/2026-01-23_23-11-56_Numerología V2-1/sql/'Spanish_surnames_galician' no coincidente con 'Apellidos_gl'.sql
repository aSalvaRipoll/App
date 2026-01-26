SELECT Spanish_surnames_galician.Apellido, Spanish_surnames_galician.Idioma, Spanish_surnames_galician.Tipo, Spanish_surnames_galician.[Raiz etimológica]
FROM Spanish_surnames_galician LEFT JOIN Apellidos_gl ON Spanish_surnames_galician.[Apellido] = Apellidos_gl.[Apellido]
WHERE (((Apellidos_gl.Apellido) Is Null));

