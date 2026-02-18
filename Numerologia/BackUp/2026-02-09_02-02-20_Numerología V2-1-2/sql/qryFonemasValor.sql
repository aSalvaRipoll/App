SELECT ID, IPA, Descripcion, Valor
FROM [Tabla fonética vocales valor]
UNION ALL
SELECT ID, IPA, Descripcion, Valor
FROM [Tabla fonética consonantes valor]
UNION ALL SELECT ID, IPA, Descripcion, Valor
FROM [Tabla fonética modificadores valor];

