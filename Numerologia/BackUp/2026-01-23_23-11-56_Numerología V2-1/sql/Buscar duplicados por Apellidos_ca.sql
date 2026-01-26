SELECT Apellidos_ca.[Apellido], Apellidos_ca.[Id]
FROM Apellidos_ca
WHERE (((Apellidos_ca.[Apellido]) In (SELECT [Apellido] FROM [Apellidos_ca] As Tmp GROUP BY [Apellido] HAVING Count(*)>1 )))
ORDER BY Apellidos_ca.[Apellido];

