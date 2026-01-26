SELECT [Apellidos_ca-ib_1].Apellido, "ca-ib" AS idioma
FROM [Apellidos_ca-ib_1] LEFT JOIN [Apellidos_ca-ib] ON [Apellidos_ca-ib_1].[Apellido] = [Apellidos_ca-ib].[Apellido]
WHERE ((([Apellidos_ca-ib].Apellido) Is Null));

