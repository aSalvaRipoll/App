CREATE TABLE tbmDicExcepciones (
    IDExcepcion AUTOINCREMENT PRIMARY KEY,
    Idioma TEXT(10) NOT NULL,
    Palabra TEXT(100) NOT NULL,
    Grafemas TEXT(255) NOT NULL,
    idFonemas TEXT(255) NOT NULL,
    EsVocal TEXT(255) NOT NULL,
    Valor TEXT(255) NOT NULL,
    Notas TEXT(255),
    Activo YESNO
);


