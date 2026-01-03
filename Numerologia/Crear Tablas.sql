CREATE TABLE tbuPersonas (
    IDPersona AUTOINCREMENT PRIMARY KEY,
    Nombre VARCHAR(150),
    IdiomaNombre BYTE,
    Ape1 VARCHAR(200),
    IdiomaApe1 BYTE,
    Ape2 VARCHAR(200),
    IdiomaApe2 BYTE,
    FechaNacimiento DATE,
    SexoLegal VARCHAR(2),
    FonNombre VARCHAR(150),
    FonApe1 VARCHAR(200),
    FonApe2 VARCHAR(200),
    FechaAlta DATE,
    FechaModificacion DATE
);


CREATE TABLE tbuResultados (
    IDResultado AUTOINCREMENT PRIMARY KEY,
    IDPersona LONG,
    FechaCalculo DATE,

    -- Números base
    NumeroDestino VARCHAR(10),
    NumeroAlma VARCHAR(10),
    NumeroPersonalidad VARCHAR(10),
    NumeroCaminoVida VARCHAR(10),
    NumeroMadurez VARCHAR(10),

    -- Números dinámicos
    AñoPersonal VARCHAR(10),
    MesPersonal VARCHAR(10),
    DiaPersonal VARCHAR(10),
    EdadPersonal VARCHAR(10),
    CicloActual BYTE,
    PinaculoActual BYTE,
    DesafioActual BYTE
);

ALTER TABLE tbuResultados
ADD CONSTRAINT FK_Resultados_Personas
FOREIGN KEY (IDPersona) REFERENCES tbuPersonas(IDPersona);




'CREATE TABLE tbuCiclos (
'    IDCiclo AUTOINCREMENT PRIMARY KEY,
'    IDResultado LONG,
'    NumeroCiclo BYTE,
'    FechaInicio DATE,
'    FechaFin DATE,
'    TextoInterpretacion LONGTEXT (Sobra, es redundante)
');

'CREATE TABLE tbuCiclos (
    IDCiclo AUTOINCREMENT PRIMARY KEY,
    IDResultado LONG,
    NumeroCiclo BYTE,
    FechaInicio DATE,
    FechaFin DATE,
);

ALTER TABLE tbuCiclos
ADD CONSTRAINT FK_Ciclos_Resultados
FOREIGN KEY (IDResultado) REFERENCES tbuResultados(IDResultado);



CREATE TABLE tbuPinaDes (
    IDPinaDes AUTOINCREMENT PRIMARY KEY,
    IDResultado LONG,

    Pina1 VARCHAR(10),
    Pina2 VARCHAR(10),
    Pina3 VARCHAR(10),
    Pina4 VARCHAR(10),

    Desa1 VARCHAR(10),
    Desa2 VARCHAR(10),
    Desa3 VARCHAR(10),
    Desa4 VARCHAR(10),

    fIni1 DATE,
    fFin1 DATE,
    fIni2 DATE,
    fFin2 DATE,
    fIni3 DATE,
    fFin3 DATE,
    fIni4 DATE,
    fFin4 DATE

);

ALTER TABLE tbuPinaDes
ADD CONSTRAINT FK_PinaDes_Resultados
FOREIGN KEY (IDResultado) REFERENCES tbuResultados(IDResultado);


CREATE TABLE tbmIdiomas (
    IDIdioma BYTE PRIMARY KEY,
    Abreviado VARCHAR(10),
    NomIdioma VARCHAR(100),
    Notas VARCHAR(255)
);

