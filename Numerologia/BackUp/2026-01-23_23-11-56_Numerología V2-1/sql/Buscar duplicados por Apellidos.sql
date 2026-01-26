SELECT First(Apellidos.[Idioma]) AS IdiomaCampo, First(Apellidos.[Apellido]) AS ApellidoCampo, Count(Apellidos.[Idioma]) AS NúmeroDeDuplicados
FROM Apellidos
GROUP BY Apellidos.[Idioma], Apellidos.[Apellido]
HAVING (((Count(Apellidos.[Idioma]))>1) AND ((Count(Apellidos.[Apellido]))>1));

