SELECT qryFonemasValor.ID, qryFonemasValor.Grafema, qryFonemasValor.IPA, qryFonemasValor.Descripcion, Len([Grafema]) AS Expr1
FROM qryFonemasValor
ORDER BY Len([Grafema]) DESC;

