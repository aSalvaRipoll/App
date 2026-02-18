SELECT ID, Grafema, IPA, Descripcion, Valor
FROM tbmVocGrafemas
UNION ALL SELECT ID, Grafema, IPA, Descripcion, Valor
FROM tbmConGrafemas;

