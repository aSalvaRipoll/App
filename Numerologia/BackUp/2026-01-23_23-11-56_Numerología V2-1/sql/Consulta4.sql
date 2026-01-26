INSERT INTO Apellidos ( Idioma, Apellido )
SELECT DISTINCT "pt-eu" AS Idioma, [Apellidos_pt-eu].Apellido
FROM [Apellidos_pt-eu];

