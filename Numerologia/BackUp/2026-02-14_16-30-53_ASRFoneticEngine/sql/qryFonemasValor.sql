SELECT ID, Grafema, IPA, Descripcion, Valor
FROM tbmVocGrafemas
UNION ALL
SELECT ID, Grafema, IPA, Descripcion, Valor
FROM tbmConGrafemas
UNION ALL SELECT ID, ASCII, IPA, Descripcion, Valor
FROM tbmModificadores;

