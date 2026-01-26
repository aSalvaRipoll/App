SELECT Apellidos.Idioma, Count(Apellidos.Apellido) AS CuentaDeApellido
FROM Apellidos
GROUP BY Apellidos.Idioma;

