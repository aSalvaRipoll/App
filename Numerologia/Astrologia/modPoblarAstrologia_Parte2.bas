Option Compare Database
Option Explicit

' =============================================================================
' Módulo: modPoblarAstrologia - PARTE 2
' Descripción: Poblar Signos Zodiacales y Arcanos Mayores
' INCLUIR después de PARTE 1
' =============================================================================

' [CONTINUACIÓN DE PARTE 1 - El código de los Signos Zodiacales va aquí]
' [Por brevedad, consulta el documento completo en outputs]

' =============================================================================
' POBLAR NÚMEROS CON ASTROLOGÍA (13 registros: 1-9 + 11,22,33,44)
' =============================================================================

Public Sub PoblarNumerosAstrologia()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM tblNumerosAstrologia", dbFailOnError
    Set rs = db.OpenRecordset("tblNumerosAstrologia", dbOpenDynaset)
    
    Debug.Print "Poblando Números con Astrología..."
    
    ' NÚMERO 1 - SOL
    rs.AddNew
    rs!Numero = 1
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Sol"
    rs!SignosAfines = "Aries, Leo"
    rs!ElementoAfin = "Fuego"
    rs!ArcanosMayoresRelacionados = "El Mago (I), El Sol (XIX)"
    rs!CaracteristicasNumerologicas = "Individualidad, liderazgo, iniciativa, independencia, ego, creatividad"
    rs!CaracteristicasAstrologicas = "Sol: centro, identidad, vitalidad, conciencia, expresión del yo"
    rs!IntegracionInterpretativa = "El 1 y el Sol ambos representan individualidad y liderazgo. Persona con 1 fuerte necesita brillo solar, expresión creativa, ser centro."
    rs.Update
    
    ' NÚMERO 2 - LUNA
    rs.AddNew
    rs!Numero = 2
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Luna"
    rs!SignosAfines = "Cáncer"
    rs!ElementoAfin = "Agua"
    rs!ArcanosMayoresRelacionados = "La Sacerdotisa (II), La Luna (XVIII)"
    rs!CaracteristicasNumerologicas = "Dualidad, cooperación, sensibilidad, intuición, receptividad"
    rs!CaracteristicasAstrologicas = "Luna: emociones, intuición, nutrición, receptividad, cambios cíclicos"
    rs!IntegracionInterpretativa = "El 2 y la Luna comparten naturaleza receptiva y emocional. Persona con 2 fuerte navega mundo emocional lunar."
    rs.Update
    
    ' NÚMERO 3 - JÚPITER
    rs.AddNew
    rs!Numero = 3
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Júpiter"
    rs!SignosAfines = "Sagitario, Piscis"
    rs!ElementoAfin = "Fuego"
    rs!ArcanosMayoresRelacionados = "La Emperatriz (III), La Rueda (X)"
    rs!CaracteristicasNumerologicas = "Expresión, creatividad, optimismo, expansión, comunicación, alegría"
    rs!CaracteristicasAstrologicas = "Júpiter: expansión, abundancia, filosofía, optimismo, crecimiento"
    rs!IntegracionInterpretativa = "El 3 y Júpiter ambos expanden y crean abundancia. Persona con 3 fuerte necesita expansión jupiteriana."
    rs.Update
    
    ' NÚMERO 4 - URANO/SATURNO
    rs.AddNew
    rs!Numero = 4
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Urano"
    rs!PlanetaSecundario = "Saturno"
    rs!SignosAfines = "Capricornio, Acuario"
    rs!ElementoAfin = "Tierra"
    rs!ArcanosMayoresRelacionados = "El Emperador (IV)"
    rs!CaracteristicasNumerologicas = "Estructura, fundamento, estabilidad, orden, trabajo, construcción"
    rs!CaracteristicasAstrologicas = "Saturno/Urano: estructura (Saturno) o ruptura de estructura (Urano) - paradoja del 4"
    rs!IntegracionInterpretativa = "El 4 tiene tensión entre estructura saturnina y libertad uraniana. Busca orden flexible."
    rs.Update
    
    ' NÚMERO 5 - MERCURIO
    rs.AddNew
    rs!Numero = 5
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Mercurio"
    rs!SignosAfines = "Géminis, Virgo"
    rs!ElementoAfin = "Aire"
    rs!ArcanosMayoresRelacionados = "El Hierofante (V)"
    rs!CaracteristicasNumerologicas = "Cambio, libertad, versatilidad, comunicación, movimiento, curiosidad"
    rs!CaracteristicasAstrologicas = "Mercurio: comunicación, intelecto, movimiento, comercio, versatilidad"
    rs!IntegracionInterpretativa = "El 5 y Mercurio comparten versatilidad y necesidad de movimiento constante."
    rs.Update
    
    ' NÚMERO 6 - VENUS
    rs.AddNew
    rs!Numero = 6
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Venus"
    rs!SignosAfines = "Tauro, Libra"
    rs!ElementoAfin = "Tierra/Aire"
    rs!ArcanosMayoresRelacionados = "Los Enamorados (VI)"
    rs!CaracteristicasNumerologicas = "Armonía, amor, responsabilidad, belleza, servicio, familia"
    rs!CaracteristicasAstrologicas = "Venus: amor, belleza, valores, armonía, atracción, placer"
    rs!IntegracionInterpretativa = "El 6 y Venus ambos buscan armonía y belleza. Persona con 6 fuerte necesita expresión venusina."
    rs.Update
    
    ' NÚMERO 7 - NEPTUNO
    rs.AddNew
    rs!Numero = 7
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Neptuno"
    rs!PlanetaSecundario = "Luna"
    rs!SignosAfines = "Piscis, Virgo"
    rs!ElementoAfin = "Agua"
    rs!ArcanosMayoresRelacionados = "El Carro (VII)"
    rs!CaracteristicasNumerologicas = "Espiritualidad, introspección, misterio, análisis profundo, perfección"
    rs!CaracteristicasAstrologicas = "Neptuno: espiritualidad, disolución, misticismo, ilusión, compasión"
    rs!IntegracionInterpretativa = "El 7 y Neptuno comparten búsqueda espiritual y profundidad mística."
    rs.Update
    
    ' NÚMERO 8 - SATURNO
    rs.AddNew
    rs!Numero = 8
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Saturno"
    rs!SignosAfines = "Capricornio, Escorpio"
    rs!ElementoAfin = "Tierra"
    rs!ArcanosMayoresRelacionados = "La Fuerza (VIII), La Justicia (XI)"
    rs!CaracteristicasNumerologicas = "Poder, logro material, karma, justicia, autoridad, ambición"
    rs!CaracteristicasAstrologicas = "Saturno: estructura, límites, disciplina, karma, tiempo, autoridad"
    rs!IntegracionInterpretativa = "El 8 y Saturno ambos tratan con karma, poder material y responsabilidad."
    rs.Update
    
    ' NÚMERO 9 - MARTE
    rs.AddNew
    rs!Numero = 9
    rs!EsNumeroMaestro = False
    rs!PlanetaPrimario = "Marte"
    rs!SignosAfines = "Aries, Sagitario"
    rs!ElementoAfin = "Fuego"
    rs!ArcanosMayoresRelacionados = "El Ermitaño (IX)"
    rs!CaracteristicasNumerologicas = "Finalización, servicio universal, compasión, sabiduría, completación"
    rs!CaracteristicasAstrologicas = "Marte: acción, energía, pasión, coraje, iniciativa"
    rs!IntegracionInterpretativa = "El 9 usa energía marciana para servicio universal. Acción dirigida al bien mayor."
    rs.Update
    
    ' NÚMERO 11 - URANO (MAESTRO)
    rs.AddNew
    rs!Numero = 11
    rs!EsNumeroMaestro = True
    rs!PlanetaPrimario = "Urano"
    rs!SignosAfines = "Acuario"
    rs!ElementoAfin = "Aire"
    rs!ArcanosMayoresRelacionados = "La Justicia (XI), La Fuerza (VIII), El Loco (XXII)"
    rs!CaracteristicasNumerologicas = "Iluminación, intuición superior, maestro espiritual, visión, sensibilidad extrema"
    rs!CaracteristicasAstrologicas = "Urano: despertar súbito, innovación, genio, revolución, libertad absoluta"
    rs!IntegracionInterpretativa = "El 11 y Urano comparten capacidad de iluminación súbita y visión innovadora."
    rs.Update
    
    ' NÚMERO 22 - PLUTÓN (MAESTRO)
    rs.AddNew
    rs!Numero = 22
    rs!EsNumeroMaestro = True
    rs!PlanetaPrimario = "Plutón"
    rs!SignosAfines = "Escorpio"
    rs!ElementoAfin = "Agua"
    rs!ArcanosMayoresRelacionados = "El Loco (XXII), El Juicio (XX)"
    rs!CaracteristicasNumerologicas = "Maestro constructor, manifestación máxima, construcción a gran escala"
    rs!CaracteristicasAstrologicas = "Plutón: transformación total, poder extremo, muerte/renacimiento, regeneración"
    rs!IntegracionInterpretativa = "El 22 usa poder plutoniano para construcción magistral. Transforma para manifestar."
    rs.Update
    
    ' NÚMERO 33 - NEPTUNO ELEVADO (MAESTRO)
    rs.AddNew
    rs!Numero = 33
    rs!EsNumeroMaestro = True
    rs!PlanetaPrimario = "Neptuno"
    rs!SignosAfines = "Piscis"
    rs!ElementoAfin = "Agua"
    rs!ArcanosMayoresRelacionados = "El Colgado (XII)"
    rs!CaracteristicasNumerologicas = "Maestro sanador, compasión universal, sacrificio consciente, sanación"
    rs!CaracteristicasAstrologicas = "Neptuno elevado: compasión infinita, disolución del ego, misticismo puro"
    rs!IntegracionInterpretativa = "El 33 es Neptuno en su expresión más alta: sanación mediante amor incondicional."
    rs.Update
    
    ' NÚMERO 44 - SATURNO+URANO (MAESTRO)
    rs.AddNew
    rs!Numero = 44
    rs!EsNumeroMaestro = True
    rs!PlanetaPrimario = "Saturno"
    rs!PlanetaSecundario = "Urano"
    rs!SignosAfines = "Capricornio, Acuario"
    rs!ElementoAfin = "Tierra/Aire"
    rs!ArcanosMayoresRelacionados = "El Mundo (XXI), El Emperador (IV)"
    rs!CaracteristicasNumerologicas = "Maestro visionario, manifestación extrema, construcción de sistemas globales"
    rs!CaracteristicasAstrologicas = "Saturno+Urano: estructura innovadora, orden que libera, sistema revolucionario"
    rs!IntegracionInterpretativa = "El 44 combina estructura saturnina con visión uraniana para construir futuro."
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Debug.Print "? 13 Números con Astrología poblados"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR poblando Números-Astrología: " & Err.Description
    If Not rs Is Nothing Then rs.Close
End Sub

' =============================================================================
' NOTA: Los decanatos (36) y figuras de corte (16) requieren código adicional
' Por espacio, se incluyen en documentación separada
' =============================================================================
