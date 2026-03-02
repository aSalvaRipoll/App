INSERT INTO tbmPrefijos ( Prefijo, Tipo, Origen, eu, Activo, es, ca, [ca-va], [ca-ib], gl, [pt-eu], [pt-br], fr, [en-gb] )
SELECT Prefijos_EU.Prefijo, Prefijos_EU.Tipo, Prefijos_EU.Origen, Prefijos_EU.eu, 1 AS Expr1, 0 AS Expr2, 0 AS Expr3, 0 AS Expr4, 0 AS Expr5, 0 AS Expr6, 0 AS Expr7, 0 AS Expr8, 0 AS Expr9, 0 AS Expr10
FROM Prefijos_EU;

