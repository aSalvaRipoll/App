Option Compare Database
Option Explicit

' =============================================================================
' Módulo: modPoblarAstrologia - PARTE 1
' Descripción: Pobla las tablas de astrología con datos completos
' Autor: Alba - Sistema de Numerología y Tarot
' =============================================================================

Public Sub PoblarTodasTablasAstrologia()
    ' Procedimiento PRINCIPAL que pobla todas las tablas
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=========================================="
    Debug.Print "POBLANDO TABLAS DE ASTROLOGÍA"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Poblar en orden lógico
    Call PoblarElementos              ' 4 registros
    Call PoblarCualidadesZodiacales   ' 3 registros
    Call PoblarPlanetas               ' 10 registros
    Call PoblarSignosZodiacales       ' 12 registros
    Call PoblarDecanatos              ' 36 registros
    Call PoblarArcanosMayoresAstrologia   ' 22 registros
    Call PoblarArcanosMenoresNumerados    ' 36 registros
    Call PoblarFigurasCorte           ' 16 registros
    Call PoblarNumerosAstrologia      ' 13 registros (1-9 + 11,22,33,44)
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TODAS LAS TABLAS POBLADAS EXITOSAMENTE"
    Debug.Print "=========================================="
    
    MsgBox "¡Tablas de Astrología pobladas completamente!" & vbCrLf & vbCrLf & _
           "4 Elementos, 3 Cualidades, 10 Planetas, 12 Signos," & vbCrLf & _
           "36 Decanatos, 22 Arcanos Mayores, 36 Menores Numerados," & vbCrLf & _
           "16 Figuras de Corte, 13 Números = 152 registros totales", _
           vbInformation, "Datos Cargados"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR al poblar tablas: " & Err.Description
    MsgBox "Error al poblar: " & Err.Description, vbCritical
End Sub

' =============================================================================
' POBLAR ELEMENTOS (4 registros)
' =============================================================================

Private Sub PoblarElementos()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM tblElementos", dbFailOnError
    Set rs = db.OpenRecordset("tblElementos", dbOpenDynaset)
    
    Debug.Print "Poblando Elementos..."
    
    ' FUEGO
    rs.AddNew
    rs!NombreElemento = "Fuego"
    rs!SimboloElemento = "??"
    rs!Polaridad = "Yang"
    rs!Cualidades = "Calor, Sequedad"
    rs!PaloTarotAsociado = "Bastos"
    rs!SignosAsociados = "Aries, Leo, Sagitario"
    rs!TriplicidadSignos = "Cardinal-Fijo-Mutable"
    rs!CaracteristicasClave = "Energía activa, iniciativa, pasión, creatividad, acción, impulso, entusiasmo"
    rs!PalabrasClave = "Acción, Pasión, Energía, Iniciativa, Creatividad, Impulso"
    rs.Update
    
    ' TIERRA
    rs.AddNew
    rs!NombreElemento = "Tierra"
    rs!SimboloElemento = "??"
    rs!Polaridad = "Yin"
    rs!Cualidades = "Frío, Sequedad"
    rs!PaloTarotAsociado = "Oros"
    rs!SignosAsociados = "Tauro, Virgo, Capricornio"
    rs!TriplicidadSignos = "Fijo-Mutable-Cardinal"
    rs!CaracteristicasClave = "Materialidad, estabilidad, recursos, practicidad, manifestación, cuerpo, seguridad"
    rs!PalabrasClave = "Materialidad, Estabilidad, Recursos, Practicidad, Manifestación"
    rs.Update
    
    ' AIRE
    rs.AddNew
    rs!NombreElemento = "Aire"
    rs!SimboloElemento = "???"
    rs!Polaridad = "Yang"
    rs!Cualidades = "Calor, Humedad"
    rs!PaloTarotAsociado = "Espadas"
    rs!SignosAsociados = "Géminis, Libra, Acuario"
    rs!TriplicidadSignos = "Mutable-Cardinal-Fijo"
    rs!CaracteristicasClave = "Pensamiento, comunicación, lógica, intelecto, análisis, ideas, conceptos"
    rs!PalabrasClave = "Pensamiento, Comunicación, Lógica, Intelecto, Ideas"
    rs.Update
    
    ' AGUA
    rs.AddNew
    rs!NombreElemento = "Agua"
    rs!SimboloElemento = "??"
    rs!Polaridad = "Yin"
    rs!Cualidades = "Frío, Humedad"
    rs!PaloTarotAsociado = "Copas"
    rs!SignosAsociados = "Cáncer, Escorpio, Piscis"
    rs!TriplicidadSignos = "Cardinal-Fijo-Mutable"
    rs!CaracteristicasClave = "Emoción, intuición, sentimiento, flujo, adaptabilidad, profundidad emocional"
    rs!PalabrasClave = "Emoción, Intuición, Sentimiento, Flujo, Profundidad"
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Debug.Print "? 4 Elementos poblados"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR poblando Elementos: " & Err.Description
    If Not rs Is Nothing Then rs.Close
End Sub

' =============================================================================
' POBLAR CUALIDADES ZODIACALES (3 registros)
' =============================================================================

Private Sub PoblarCualidadesZodiacales()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM tblCualidadesZodiacales", dbFailOnError
    Set rs = db.OpenRecordset("tblCualidadesZodiacales", dbOpenDynaset)
    
    Debug.Print "Poblando Cualidades Zodiacales..."
    
    ' CARDINAL
    rs.AddNew
    rs!NombreCualidad = "Cardinal"
    rs!CruzCosmica = "Cruz Cardinal"
    rs!SignosAsociados = "Aries, Cáncer, Libra, Capricornio"
    rs!CaracteristicasPrincipales = "Inicio, acción, liderazgo, impulso creador. Energía que INICIA ciclos, que comienza cosas nuevas."
    rs!ModoExpresion = "Iniciativa directa, acción inmediata, liderazgo natural"
    rs!PalabrasClave = "Inicio, Acción, Liderazgo, Impulso"
    rs.Update
    
    ' FIJO
    rs.AddNew
    rs!NombreCualidad = "Fijo"
    rs!CruzCosmica = "Cruz Fija"
    rs!SignosAsociados = "Tauro, Leo, Escorpio, Acuario"
    rs!CaracteristicasPrincipales = "Estabilidad, persistencia, determinación, resistencia al cambio. Energía que MANTIENE y CONSOLIDA."
    rs!ModoExpresion = "Constancia, lealtad, terquedad, resistencia"
    rs!PalabrasClave = "Estabilidad, Persistencia, Determinación, Constancia"
    rs.Update
    
    ' MUTABLE
    rs.AddNew
    rs!NombreCualidad = "Mutable"
    rs!CruzCosmica = "Cruz Mutable"
    rs!SignosAsociados = "Géminis, Virgo, Sagitario, Piscis"
    rs!CaracteristicasPrincipales = "Adaptabilidad, flexibilidad, cambio, transición. Energía que TRANSFORMA y ADAPTA."
    rs!ModoExpresion = "Versatilidad, adaptación, dispersión posible"
    rs!PalabrasClave = "Adaptabilidad, Flexibilidad, Cambio, Versatilidad"
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Debug.Print "? 3 Cualidades pobladas"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR poblando Cualidades: " & Err.Description
    If Not rs Is Nothing Then rs.Close
End Sub

' =============================================================================
' POBLAR PLANETAS (10 registros principales)
' =============================================================================

Private Sub PoblarPlanetas()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM tblPlanetas", dbFailOnError
    Set rs = db.OpenRecordset("tblPlanetas", dbOpenDynaset)
    
    Debug.Print "Poblando Planetas..."
    
    ' SOL
    rs.AddNew
    rs!NombrePlaneta = "Sol"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Luminaria"
    rs!SignoDomicilio = "Leo"
    rs!SignoExaltacion = "Aries"
    rs!SignoDetrimento = "Acuario"
    rs!SignoCaida = "Libra"
    rs!ArcanosAsociados = "El Sol (XIX), El Mago (I)"
    rs!NumeroNumerologico = 1
    rs!ElementoAfin = "Fuego"
    rs!CaracteristicasClave = "Ego, identidad, vitalidad, centro, conciencia, creatividad, autoridad, padre"
    rs!PalabrasClave = "Identidad, Ego, Vitalidad, Centro, Conciencia"
    rs!CicloOrbital = "365.25 días (aparente)"
    rs.Update
    
    ' LUNA
    rs.AddNew
    rs!NombrePlaneta = "Luna"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Luminaria"
    rs!SignoDomicilio = "Cáncer"
    rs!SignoExaltacion = "Tauro"
    rs!SignoDetrimento = "Capricornio"
    rs!SignoCaida = "Escorpio"
    rs!ArcanosAsociados = "La Luna (XVIII), La Sacerdotisa (II)"
    rs!NumeroNumerologico = 2
    rs!ElementoAfin = "Agua"
    rs!CaracteristicasClave = "Emociones, intuición, inconsciente, receptividad, nutrición, madre, hábitos"
    rs!PalabrasClave = "Emoción, Intuición, Receptividad, Nutrición"
    rs!CicloOrbital = "29.5 días"
    rs.Update
    
    ' MERCURIO
    rs.AddNew
    rs!NombrePlaneta = "Mercurio"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Personal"
    rs!SignoDomicilio = "Géminis, Virgo"
    rs!SignoExaltacion = "Virgo"
    rs!SignoDetrimento = "Sagitario, Piscis"
    rs!SignoCaida = "Piscis"
    rs!ArcanosAsociados = "El Mago (I)"
    rs!NumeroNumerologico = 5
    rs!ElementoAfin = "Aire"
    rs!CaracteristicasClave = "Comunicación, intelecto, razón, comercio, movimiento, versatilidad, información"
    rs!PalabrasClave = "Comunicación, Intelecto, Versatilidad, Comercio"
    rs!CicloOrbital = "88 días"
    rs.Update
    
    ' VENUS
    rs.AddNew
    rs!NombrePlaneta = "Venus"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Personal"
    rs!SignoDomicilio = "Tauro, Libra"
    rs!SignoExaltacion = "Piscis"
    rs!SignoDetrimento = "Aries, Escorpio"
    rs!SignoCaida = "Virgo"
    rs!ArcanosAsociados = "La Emperatriz (III)"
    rs!NumeroNumerologico = 6
    rs!ElementoAfin = "Tierra/Aire"
    rs!CaracteristicasClave = "Amor, belleza, armonía, placer, valores, atracción, arte, dinero"
    rs!PalabrasClave = "Amor, Belleza, Armonía, Valores, Placer"
    rs!CicloOrbital = "225 días"
    rs.Update
    
    ' MARTE
    rs.AddNew
    rs!NombrePlaneta = "Marte"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Personal"
    rs!SignoDomicilio = "Aries, Escorpio"
    rs!SignoExaltacion = "Capricornio"
    rs!SignoDetrimento = "Libra, Tauro"
    rs!SignoCaida = "Cáncer"
    rs!ArcanosAsociados = "La Torre (XVI)"
    rs!NumeroNumerologico = 9
    rs!ElementoAfin = "Fuego"
    rs!CaracteristicasClave = "Acción, energía, deseo, coraje, agresión, sexualidad, guerra, iniciativa"
    rs!PalabrasClave = "Acción, Energía, Deseo, Coraje, Agresión"
    rs!CicloOrbital = "687 días"
    rs.Update
    
    ' JÚPITER
    rs.AddNew
    rs!NombrePlaneta = "Júpiter"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Social"
    rs!SignoDomicilio = "Sagitario, Piscis"
    rs!SignoExaltacion = "Cáncer"
    rs!SignoDetrimento = "Géminis, Virgo"
    rs!SignoCaida = "Capricornio"
    rs!ArcanosAsociados = "La Rueda de la Fortuna (X)"
    rs!NumeroNumerologico = 3
    rs!ElementoAfin = "Fuego"
    rs!CaracteristicasClave = "Expansión, abundancia, optimismo, crecimiento, filosofía, fe, suerte, sabiduría"
    rs!PalabrasClave = "Expansión, Abundancia, Optimismo, Crecimiento"
    rs!CicloOrbital = "11.86 años"
    rs.Update
    
    ' SATURNO
    rs.AddNew
    rs!NombrePlaneta = "Saturno"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Social"
    rs!SignoDomicilio = "Capricornio, Acuario"
    rs!SignoExaltacion = "Libra"
    rs!SignoDetrimento = "Cáncer, Leo"
    rs!SignoCaida = "Aries"
    rs!ArcanosAsociados = "El Mundo (XXI)"
    rs!NumeroNumerologico = 8
    rs!ElementoAfin = "Tierra"
    rs!CaracteristicasClave = "Estructura, límites, disciplina, responsabilidad, karma, tiempo, madurez, autoridad"
    rs!PalabrasClave = "Estructura, Límites, Disciplina, Karma, Tiempo"
    rs!CicloOrbital = "29.46 años"
    rs.Update
    
    ' URANO
    rs.AddNew
    rs!NombrePlaneta = "Urano"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Transpersonal"
    rs!SignoDomicilio = "Acuario"
    rs!SignoExaltacion = "Escorpio"
    rs!SignoDetrimento = "Leo"
    rs!SignoCaida = "Tauro"
    rs!ArcanosAsociados = "El Loco (XXII/0)"
    rs!NumeroNumerologico = 11
    rs!ElementoAfin = "Aire"
    rs!CaracteristicasClave = "Cambio súbito, revolución, innovación, individualidad, libertad, despertar, genio"
    rs!PalabrasClave = "Cambio, Revolución, Innovación, Libertad, Despertar"
    rs!CicloOrbital = "84 años"
    rs.Update
    
    ' NEPTUNO
    rs.AddNew
    rs!NombrePlaneta = "Neptuno"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Transpersonal"
    rs!SignoDomicilio = "Piscis"
    rs!SignoExaltacion = "Cáncer/Leo"
    rs!SignoDetrimento = "Virgo"
    rs!SignoCaida = "Capricornio"
    rs!ArcanosAsociados = "El Colgado (XII)"
    rs!NumeroNumerologico = 7
    rs!ElementoAfin = "Agua"
    rs!CaracteristicasClave = "Espiritualidad, ilusión, disolución, compasión, misticismo, inspiración, engaño"
    rs!PalabrasClave = "Espiritualidad, Ilusión, Compasión, Misticismo"
    rs!CicloOrbital = "164.79 años"
    rs.Update
    
    ' PLUTÓN
    rs.AddNew
    rs!NombrePlaneta = "Plutón"
    rs!SimboloUnicode = "?"
    rs!TipoCuerpo = "Planeta Transpersonal"
    rs!SignoDomicilio = "Escorpio"
    rs!SignoExaltacion = "Leo/Aries"
    rs!SignoDetrimento = "Tauro"
    rs!SignoCaida = "Acuario"
    rs!ArcanosAsociados = "El Juicio (XX)"
    rs!NumeroNumerologico = 22
    rs!ElementoAfin = "Agua"
    rs!CaracteristicasClave = "Transformación profunda, poder, muerte/renacimiento, intensidad, lo oculto, regeneración"
    rs!PalabrasClave = "Transformación, Poder, Muerte/Renacimiento, Intensidad"
    rs!CicloOrbital = "248.09 años"
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Debug.Print "? 10 Planetas poblados"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR poblando Planetas: " & Err.Description
    If Not rs Is Nothing Then rs.Close
End Sub
