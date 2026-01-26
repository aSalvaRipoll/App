CREATE TABLE tbmDicExcepciones (
    IDExcepcion AUTOINCREMENT PRIMARY KEY,
    Idioma TEXT(10),
    Tipo TEXT(10),
    Palabra TEXT(100),
    Grafema TEXT(10),
    FonemaCompleto TEXT(255),
    idFonema BYTE,
    Notas TEXT(255),
    Activo YESNO
);


