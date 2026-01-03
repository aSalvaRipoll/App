Type ProgAnual
    Fisico As String * 1
    Mental As String * 1
    Espiritual As String * 1
    EsenciaDD As Byte
    EsenciaSD As Byte
End Type

Type SalidaDatos
    Vocales As String
    Consonantes As String
    Completo As String
End Type

Type Resultados
    Cadena As String
    Inicial As Byte
    Medio As Byte
    Final As Byte
    Maestro As Byte
    Karma As Byte
End Type

Type Prim
    vocal As String
    Conso As String
End Type

Type CalcPersonal
    PAño As Integer
    PMes As Integer
    PDia As Integer
    PSem As Integer
    Ciclo As Byte
    ValorCiclo As Integer
End Type

Type DatosFecha
    FechaNacimiento As Date
    Dia As Byte
    Mes As Byte
    Signo As String
    
    EdadPer As Byte
    ResAño As Integer
    Sendero As Resultados
    ResAñoP As Byte
    AñoP As Resultados
    MesP As Resultados
    SemP As Resultados
    DiaP As Resultados
    Natalicio As Integer
    Apoyo As Resultados

    PinaEscoM As Boolean
    PinaEsco(1 To 2, 1 To 4) As Resultados
    Pina(1 To 4) As Resultados
    Esco(1 To 4) As Resultados

    Edades(1 To 4) As Byte
    EdadesM(1 To 4) As Byte

    Periodicos(1 To 3) As Resultados
    FechasPeriodicos(1 To 3) As String

    V_Temp As Resultados
    V_Temporal As Byte
    V_TemporalR As Byte

    SubDes1 As Integer
    SubDes2 As Integer
    SubDes3 As Integer
End Type

Type DatosPersona

    NomP As String
    ape1 As String
    ape2 As String

    NombrePersona As String
    CadenaNum As SalidaDatos
    Persona As Boolean

    CPares As Byte
    CImpar As Byte

    Alma As Resultados
    Camino As Resultados
    PExt As Resultados

    Apoyo As Resultados
    Evolucion As Resultados

    Casas(1 To 9) As Byte
    Habitantes(1 To 9) As Byte
    Puentes(1 To 9) As Byte
    Evoluciones(1 To 9) As Byte
    EsEvolucion(1 To 9) As Byte
    Inconscientes(1 To 9) As Byte
    
    Subconsciente As Byte
    
    Porcen(1 To 9, 3) As Double

    Planos(1 To 4) As Byte

    Ciclos(1 To 3) As String

    Pasion() As Byte

    Paterna As Resultados
    Materna As Resultados
    Herencia As Resultados

    Madurez As Resultados
    Equilibrio As Resultados
    PensamientoRacional As Resultados
    YoSubcons As Resultados

    Transito As ProgAnual
End Type

Public PerFecha As DatosFecha
Public PerDatos As DatosPersona