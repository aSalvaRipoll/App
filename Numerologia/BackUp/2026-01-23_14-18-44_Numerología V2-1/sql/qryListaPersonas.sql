SELECT tbuPersonas.ID_Persona, Nombre & ' ' & Ape1 & ' ' & Ape2 AS NombreCompleto, tbuPersonas.FechaNacimiento, tbmGeneros.Genero
FROM tbuPersonas INNER JOIN tbmGeneros ON tbuPersonas.[ID_Genero] = tbmGeneros.ID
ORDER BY tbuPersonas.Nombre, tbuPersonas.Ape1, tbuPersonas.Ape2;

