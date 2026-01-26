SELECT [Apellidos_ca-ib].[Apellido], [Apellidos_ca-ib].[Id]
FROM [Apellidos_ca-ib]
WHERE ((([Apellidos_ca-ib].[Apellido]) In (SELECT [Apellido] FROM [Apellidos_ca-ib] As Tmp GROUP BY [Apellido] HAVING Count(*)>1 )))
ORDER BY [Apellidos_ca-ib].[Apellido];

