SELECT qryFonemasValor.ID, First(qryFonemasValor.IPA) AS PrimeroDeIPA
FROM qryFonemasValor
GROUP BY qryFonemasValor.ID
ORDER BY qryFonemasValor.ID, First(qryFonemasValor.IPA);

