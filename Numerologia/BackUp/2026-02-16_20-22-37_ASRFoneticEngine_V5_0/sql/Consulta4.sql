INSERT INTO tbmPrefijos ( Prefijo, Tipo, Activo )
SELECT Prefijos.Prefijo, Prefijos.Tipo, IIf([Activo] Like "true",1,0) AS Expr1
FROM Prefijos;

