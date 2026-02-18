INSERT INTO tbmApellidos ( Idioma, Apellido )
SELECT DISTINCT Apellidos.Idioma, Apellidos.Apellido
FROM Apellidos;

