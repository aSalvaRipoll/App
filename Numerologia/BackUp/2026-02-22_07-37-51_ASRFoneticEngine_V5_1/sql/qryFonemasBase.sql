SELECT ID, Grafema, IPA, Descripcion, Valor, 'V' as Tipo
FROM tbmVocGrafemas
UNION ALL SELECT ID, Grafema, IPA, Descripcion, Valor, 'C' as Tipo
FROM tbmConGrafemas;

