INSERT INTO tbmNombres ( nombre, idioma, genero, notas, Activo )
SELECT DISTINCT [Nombres EN-US].nombre, [Nombres EN-US].idioma, [Nombres EN-US].genero, [Nombres EN-US].notas, True AS Expr1
FROM [Nombres EN-US]
WHERE ((([Nombres EN-US].nombre) Is Not Null));

