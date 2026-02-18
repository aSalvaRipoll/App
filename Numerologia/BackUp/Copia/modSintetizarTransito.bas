Attribute VB_Name = "modSintetizarTransito"
' ------------------------------------------------------
' Nombre:    modSintetizarTransito
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------
Option Compare Database
Option Explicit


Public Function BuscarCorrespondencias(ByVal Numero As Long) As DAO.Recordset

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_NumerologiaAstrologia WHERE ID = " & Numero, _
        dbOpenSnapshot)

    If rs.EOF Then
        Set BuscarCorrespondencias = Nothing
    Else
        Set BuscarCorrespondencias = rs
    End If

End Function



'Function SintetizarTránsito_1(ID As Long, Plano As String) As String
'
'    Dim rs As DAO.Recordset
'    Set rs = BuscarCorrespondencias(ID)
'
'    If rs Is Nothing Then
'        SintetizarTránsito = "Sin datos."
'        Exit Function
'    End If
'
'    Dim sElemento As String
'    Dim sModalidad As String
'    Dim sSigno As String
'    Dim sArcano As String
'    Dim sDecanato As String
'    Dim sPlano As String
'
'    sElemento = InterpretarElemento(rs!Elemento)
'    sModalidad = InterpretarModalidad(rs!Modalidad)
'    sSigno = InterpretarSigno(rs!Signo)
'    sArcano = InterpretarArcano(rs!Arcano)
'    sDecanato = InterpretarDecanato(rs!Decanato)
'    sPlano = InterpretarPlano(Plano)
'
'    SintetizarTránsito = sPlano & ": " & sArcano & ", " & sSigno & " " & sDecanato & _
'                         ", con energía de " & sElemento & " y dinámica " & sModalidad & "."
'
'End Function
'
'Function SintetizarTránsito_2(ID As Long, Plano As String) As String
'
'    Dim rs As DAO.Recordset
'    Set rs = BuscarCorrespondencias(ID)
'
'    If rs Is Nothing Then
'        SintetizarTránsito = "Sin datos."
'        Exit Function
'    End If
'
'    Dim sPlano As String
'    Dim sElemento As String
'    Dim sModalidad As String
'    Dim sSigno As String
'    Dim sDecanato As String
'    Dim sArcano As String
'    Dim sFigura As String
'    Dim sNumero As String
'    Dim TipoArcano As String
'
'    sPlano = InterpretarPlano(Plano)
'    sElemento = InterpretarElemento(rs!Elemento)
'    sModalidad = InterpretarModalidad(rs!Modalidad)
'    sSigno = InterpretarSigno(rs!Signo)
'    sDecanato = InterpretarDecanato(rs!Decanato)
'
'    TipoArcano = rs!TipoArcano
'
'    ' Detectar figura
'    sFigura = Split(rs!Arcano, " ")(0)
'
'    Select Case TipoArcano
'
'        Case "Mayor"
'            sArcano = InterpretarArcano(rs!Arcano)
'
'        Case "Corte"
'            sArcano = InterpretarFigura(sFigura)
'
'        Case "As"
'            sArcano = InterpretarNumero(1)
'
'        Case "Menor"
'            sNumero = rs!Reduccion
'            sArcano = InterpretarNumero(sNumero)
'
'        Case Else
'            sArcano = ""
'    End Select
'
'    SintetizarTránsito = sPlano & ": " & sArcano & ", con energía de " & sElemento & _
'                         ", dinámica " & sModalidad & ", tono " & sSigno & _
'                         " y " & sDecanato & "."
'
'End Function
'
'Function SintetizarTránsito_3(ID As Long, Plano As String) As String
'
'    Dim rs As DAO.Recordset
'    Set rs = BuscarCorrespondencias(ID)
'
'    If rs Is Nothing Then
'        SintetizarTránsito = "Sin datos."
'        Exit Function
'    End If
'
'    Dim sPlano As String
'    Dim sElemento As String
'    Dim sModalidad As String
'    Dim sSigno As String
'    Dim sDecanato As String
'    Dim sArcano As String
'    Dim sFigura As String
'    Dim sNumero As String
'    Dim sPalo As String
'    Dim TipoArcano As String
'
'    sPlano = InterpretarPlano(Plano)
'    sElemento = InterpretarElemento(rs!Elemento)
'    sModalidad = InterpretarModalidad(rs!Modalidad)
'    sSigno = InterpretarSigno(rs!Signo)
'    sDecanato = InterpretarDecanato(rs!Decanato)
'    sPalo = InterpretarPalo(rs!Palo)
'
'    TipoArcano = rs!TipoArcano
'    sFigura = Split(rs!Arcano, " ")(0)
'
'    Select Case TipoArcano
'
'        Case "Mayor"
'            sArcano = InterpretarArcano(rs!Arcano)
'
'        Case "Corte"
'            sArcano = InterpretarFigura(sFigura)
'
'        Case "As"
'            sArcano = InterpretarNumero(1)
'
'        Case "Menor"
'            sNumero = rs!Reduccion
'            sArcano = InterpretarNumero(sNumero)
'
'        Case Else
'            sArcano = ""
'    End Select
'
'    SintetizarTránsito = sPlano & ": " & sArcano & ", " & sPalo & _
'                         ", con energía de " & sElemento & _
'                         ", dinámica " & sModalidad & ", tono " & sSigno & _
'                         " y " & sDecanato & "."
'
'End Function

Function SintetizarTransito(id As Long, Plano As String, Estilo As Integer) As String

    Dim rs As DAO.Recordset
    Set rs = BuscarCorrespondencias(id)

    If rs Is Nothing Then
        SintetizarTransito = "Sin datos."
        Exit Function
    End If

    Dim sPlano As String
    Dim sElemento As String
    Dim sModalidad As String
    Dim sSigno As String
    Dim sDecanato As String
    Dim sArcano As String
    Dim sFigura As String
    Dim sNumero As Integer 'String
    Dim sPalo As String
    Dim TipoArcano As String

    sPlano = InterpretarPlano(Plano)
    sElemento = InterpretarElemento(rs!Elemento)
    sModalidad = InterpretarModalidad(rs!Modalidad)
    sSigno = InterpretarSigno(rs!Signo)
    sDecanato = InterpretarDecanato(rs!Decanato)
    sPalo = InterpretarPalo(rs!Palo)

    TipoArcano = rs!TipoArcano
    sFigura = Split(rs!Arcano, " ")(0)

    Select Case TipoArcano

        Case "Mayor"
            sArcano = InterpretarArcano(rs!Arcano)

        Case "Corte"
            sArcano = InterpretarFigura(sFigura)

        Case "As"
            sArcano = InterpretarNumero(1)

        Case "Menor"
            sNumero = rs!Reduccion
            sArcano = InterpretarNumero(sNumero)

        Case Else
            sArcano = ""
    End Select
    
    ' Llamada al estilo seleccionado
    SintetizarTransito = EstiloTransito(Estilo, sPlano, sArcano, sPalo, _
                        sElemento, sModalidad, _
                        sSigno, sDecanato)
    
    
End Function

Function EstiloTransito(Estilo As Integer, _
                        sPlano As String, sArcano As String, sPalo As String, _
                        sElemento As String, sModalidad As String, _
                        sSigno As String, sDecanato As String) As String

    Select Case Estilo
        Case 0   ' Compacta
            EstiloTransito = sPlano & ": " & sArcano & ". " & _
                              sPalo & ", " & sElemento & ", " & _
                              sModalidad & ", " & sSigno & ", " & sDecanato & "."

        Case 1 ' Base
            EstiloTransito = sPlano & ": " & sArcano & ", " & sPalo & _
                         ", con energía de " & sElemento & _
                         ", dinámica " & sModalidad & ", tono " & sSigno & _
                         " y " & sDecanato & "."

        Case 2 ' A (técnico-elegante):
            EstiloTransito = sPlano & ". " & _
                        "Arquetipo: " & sArcano & ". " & _
                        "Palo: " & sPalo & ". " & _
                        "Elemento: " & sElemento & ". " & _
                        "Modalidad: " & sModalidad & ". " & _
                        "Signo: " & sSigno & ". " & _
                        "Decanato: " & sDecanato & "."
    
        Case 3 ' B (poético-hermético):
            EstiloTransito = sPlano & ", donde " & sArcano & _
                        " se expresa a través de " & sPalo & _
                        ", guiado por el " & sElemento & _
                        " en dinámica " & sModalidad & _
                        ", bajo el tono de " & sSigno & _
                        " y el matiz del " & sDecanato & "."



        Case 4 ' C (sintético-directo):
            EstiloTransito = sPlano & ": " & sArcano & ". " & _
                        sPalo & ". " & sElemento & ". " & _
                        sModalidad & ". " & sSigno & ". " & _
                        sDecanato & "."

        Case 5 ' Estilo híbrido: técnico + poético + sintético
            EstiloTransito = sPlano & ", esta energía se manifiesta como " & sArcano & _
                        ". Se expresa a través de " & sPalo & _
                        ", guiada por el " & sElemento & _
                        " en dinámica " & sModalidad & _
                        ", bajo el tono de " & sSigno & _
                        " y el matiz del " & sDecanato & "."

        Case 6 ' Versión compacta (rápida, directa, sintética)
            EstiloTransito = sPlano & ": " & _
                    sArcano & ". " & _
                    sPalo & ", " & _
                    sElemento & ", " & _
                    sModalidad & ", " & _
                    sSigno & ", " & _
                    sDecanato & "."


        Case 7 ' Versión extendida (híbrida, elegante, descriptiva)
            EstiloTransito = sPlano & _
                    ", esta energía se manifiesta como " & sArcano & _
                    ". Se expresa a través de " & sPalo & _
                    ", guiada por el " & sElemento & _
                    " en dinámica " & sModalidad & _
                    ", bajo el tono de " & sSigno & _
                    " y el matiz del " & sDecanato & _
                    ". Esta combinación aporta un movimiento particular que influye en el desarrollo del tránsito."

        Case 8 ' Construcción narrativa
            EstiloTransito = sPlano & _
                    ", esta energía toma forma a través de " & sArcano & _
                    ". Su expresión se despliega mediante " & sPalo & _
                    ", un cauce que revela la cualidad esencial del tránsito. " & _
                    "El " & sElemento & _
                    " aporta la tonalidad profunda, mientras que la " & sModalidad & _
                    " define el modo en que esta fuerza se mueve y se manifiesta. " & _
                    "El sello de " & sSigno & _
                    " orienta la dirección interna del proceso, y el " & sDecanato & _
                    " añade un matiz particular que colorea la vivencia. " & _
                    "En conjunto, esta combinación describe un momento donde las energías convergen para abrir un camino de comprensión y transformación."
    
        Case 9 ' Versión Profunda — la voz más contemplativa de tu sistema
            EstiloTransito = sPlano & _
                    ", esta energía se despliega como " & sArcano & _
                    ", invitando a una experiencia que nace desde un lugar íntimo y silencioso. " & _
                    "El cauce de " & sPalo & _
                    " actúa como un puente entre lo interno y lo manifestado, permitiendo que la vibración esencial encuentre su forma. " & _
                    "El " & sElemento & _
                    " aporta la raíz simbólica que sostiene el proceso, mientras que la " & sModalidad & _
                    " define el modo en que esta fuerza se mueve, respira y se transforma. " & _
                    "El sello de " & sSigno & _
                    " orienta la dirección profunda del tránsito, revelando la cualidad que guía la vivencia. " & _
                    "El " & sDecanato & _
                    " añade un matiz sutil, casi imperceptible, pero decisivo en la manera en que esta energía se encarna. " & _
                    "En conjunto, esta configuración describe un momento en el que las capas visibles e invisibles convergen, " & _
                    "abriendo un espacio de comprensión que trasciende lo inmediato y permite acceder a un nivel más hondo de significado."

       
        Case 10 ' Extendida
            EstiloTransito = sPlano & ", esta energía se manifiesta como " & sArcano & _
                              ". Se expresa a través de " & sPalo & _
                              ", guiada por el " & sElemento & _
                              " en dinámica " & sModalidad & _
                              ", bajo el tono de " & sSigno & _
                              " y el matiz del " & sDecanato & "."

        Case 11 ' Narrativa
            EstiloTransito = sPlano & ", esta energía toma forma a través de " & sArcano & _
                              ". Su expresión se despliega mediante " & sPalo & _
                              ", mientras el " & sElemento & _
                              " aporta la tonalidad profunda y la " & sModalidad & _
                              " define su movimiento. El sello de " & sSigno & _
                              " orienta el proceso, y el " & sDecanato & _
                              " añade un matiz particular."

        Case 12 ' Profunda
            EstiloTransito = sPlano & ", esta energía se despliega como " & sArcano & _
                              ", invitando a una experiencia que nace desde un lugar íntimo. " & _
                              "El cauce de " & sPalo & _
                              " actúa como puente entre lo interno y lo manifestado. " & _
                              "El " & sElemento & " sostiene el proceso, la " & sModalidad & _
                              " define su respiración, el sello de " & sSigno & _
                              " guía la vivencia y el " & sDecanato & _
                              " aporta un matiz sutil que colorea la experiencia."

        
        Case 13 ' Modo Zen
            EstiloTransito = sPlano & ". " & sArcano & _
                              ". Nada sobra. Nada falta. " & sElemento & _
                              ". " & sModalidad & ". " & sSigno & ". " & sDecanato & "."

        Case 14 ' Modo Oracular
            EstiloTransito = "En " & LCase(sPlano) & ", el signo es claro: " & sArcano & _
                              ". El elemento " & sElemento & _
                              " abre el camino, la modalidad " & sModalidad & _
                              " marca el ritmo, y " & sSigno & _
                              " revela la dirección. El " & sDecanato & _
                              " susurra el mensaje final."

        Case 15 ' Modo Pitágoras en trance
            EstiloTransito = sPlano & ": " & sArcano & _
                              " vibra en resonancia con el " & sElemento & _
                              ", modulándose a través de la " & sModalidad & _
                              ". " & sSigno & " actúa como vector, y el " & sDecanato & _
                              " completa la proporción."

        Case 16 ' Modo Vienna Philharmonic
            EstiloTransito = sPlano & ", esta energía entra como un motivo de " & sArcano & _
                              " sobre un fondo de " & sElemento & _
                              ". La " & sModalidad & _
                              " marca el compás, " & sSigno & _
                              " define la melodía, y el " & sDecanato & _
                              " aporta el color orquestal."

        Case 17 ' Modo Kaos controlado
            EstiloTransito = sPlano & ". " & sArcano & _
                              " irrumpe, el " & sElemento & _
                              " se agita, la " & sModalidad & _
                              " se retuerce, " & sSigno & _
                              " se expande y el " & sDecanato & _
                              " remata la jugada."

        Case 18 ' Modo Legacy 1992
            EstiloTransito = sPlano & ": " & sArcano & _
                              ". Elemento=" & sElemento & _
                              "; Modalidad=" & sModalidad & _
                              "; Signo=" & sSigno & _
                              "; Decanato=" & sDecanato & "."

        Case 19 ' Modo mi visión
            EstiloTransito = sPlano & ", según veo, esto es claramente " & sArcano & _
                              " con " & sElemento & _
                              ", " & sModalidad & ", " & sSigno & _
                              " y un " & sDecanato & " que no falla."

        Case Else
            EstiloTransito = "Estilo no reconocido."

    End Select

End Function


Function InterpretarElemento(Elemento As String) As String

    Select Case Elemento
                                        'Esencia, Movimiento, Luz, Sombra, Función

        Case "Fuego":   InterpretarElemento = "impulso, dirección y propósito"

        Case "Agua":    InterpretarElemento = "sensibilidad, memoria y profundidad emocional"

        Case "Aire":    InterpretarElemento = "claridad mental, comunicación y análisis"

        Case "Tierra":  InterpretarElemento = "estabilidad, forma y concreción"

        Case "Éter":    InterpretarElemento = "potencial puro y energía no manifestada"

        Case Else:
            InterpretarElemento = ""
    End Select

End Function

Function InterpretarModalidad(Modalidad As String) As String

    Select Case Modalidad

        Case "Cardinal"
            InterpretarModalidad = "inicio, impulso y dirección"

        Case "Fijo"
            InterpretarModalidad = "estabilidad, concentración y permanencia"

        Case "Mutable"
            InterpretarModalidad = "adaptación, transición y flexibilidad"

        Case "Libre"
            InterpretarModalidad = "potencial abierto y movimiento sin estructura"

        Case Else
            InterpretarModalidad = ""
    End Select

End Function

Function InterpretarSigno(Signo As String) As String

    Select Case Signo

        Case "Aries"
            InterpretarSigno = "impulso, afirmación y acción directa"

        Case "Tauro"
            InterpretarSigno = "estabilidad, permanencia y disfrute sensorial"

        Case "Géminis"
            InterpretarSigno = "curiosidad, comunicación y flexibilidad mental"

        Case "Cáncer"
            InterpretarSigno = "sensibilidad, protección y memoria emocional"

        Case "Leo"
            InterpretarSigno = "expresión, creatividad y fuerza vital"

        Case "Virgo"
            InterpretarSigno = "análisis, orden y mejora continua"

        Case "Libra"
            InterpretarSigno = "equilibrio, relación y armonía"

        Case "Escorpión"
            InterpretarSigno = "intensidad, transformación y profundidad emocional"

        Case "Sagitario"
            InterpretarSigno = "expansión, búsqueda y visión"

        Case "Capricornio"
            InterpretarSigno = "estructura, responsabilidad y logro"

        Case "Acuario"
            InterpretarSigno = "innovación, libertad y visión colectiva"

        Case "Piscis"
            InterpretarSigno = "sensibilidad, unión y trascendencia"

        Case Else
            InterpretarSigno = ""
    End Select

End Function

Function InterpretarArcano(Arcano As String) As String

    Select Case Arcano

        Case "El Loco"
            InterpretarArcano = "potencial puro y libertad"

        Case "El Mago"
            InterpretarArcano = "acción consciente y manifestación"

        Case "La Papisa"
            InterpretarArcano = "intuición y gestación interna"

        Case "La Emperatriz"
            InterpretarArcano = "creatividad y expansión natural"

        Case "El Emperador"
            InterpretarArcano = "estructura y autoridad"

        Case "El Hierofante"
            InterpretarArcano = "tradición y guía espiritual"

        Case "Los Enamorados"
            InterpretarArcano = "elección e integración"

        Case "El Carro"
            InterpretarArcano = "dirección y avance"

        Case "La Fuerza"
            InterpretarArcano = "coraje y dominio interno"

        Case "El Ermitaño"
            InterpretarArcano = "introspección y sabiduría"

        Case "La Rueda de la Fortuna"
            InterpretarArcano = "cambio y oportunidad"

        Case "La Justicia"
            InterpretarArcano = "equilibrio y claridad ética"

        Case "El Ahorcado"
            InterpretarArcano = "pausa y cambio de perspectiva"

        Case "La Muerte"
            InterpretarArcano = "transformación y renacimiento"

        Case "La Templanza"
            InterpretarArcano = "integración y armonía"

        Case "El Diablo"
            InterpretarArcano = "deseo y sombra"

        Case "La Torre"
            InterpretarArcano = "ruptura y liberación"

        Case "La Estrella"
            InterpretarArcano = "inspiración y guía"

        Case "La Luna"
            InterpretarArcano = "intuición e inconsciente"

        Case "El Sol"
            InterpretarArcano = "claridad y vitalidad"

        Case "El Juicio"
            InterpretarArcano = "despertar y renovación"

        Case "El Mundo"
            InterpretarArcano = "plenitud y cierre"

        Case Else
            InterpretarArcano = ""
    End Select

End Function

Function InterpretarNumero(Numero As Integer) As String

    Select Case Numero

        Case 1
            InterpretarNumero = "inicio y potencial del elemento"

        Case 2
            InterpretarNumero = "dualidad y equilibrio inicial"

        Case 3
            InterpretarNumero = "expansión y creatividad"

        Case 4
            InterpretarNumero = "estructura y estabilidad"

        Case 5
            InterpretarNumero = "cambio y desafío"

        Case 6
            InterpretarNumero = "armonía e integración"

        Case 7
            InterpretarNumero = "prueba y ajuste interno"

        Case 8
            InterpretarNumero = "movimiento y transformación"

        Case 9
            InterpretarNumero = "culminación y madurez"

        Case 10
            InterpretarNumero = "cierre y transición de ciclo"

        Case Else
            InterpretarNumero = ""
    End Select

End Function

Function InterpretarFigura(Figura As String) As String

    Select Case Figura

        Case "Rey"
            InterpretarFigura = "dominio, madurez y dirección consciente"

        Case "Reina"
            InterpretarFigura = "receptividad, influencia interna y sensibilidad"

        Case "Caballero"
            InterpretarFigura = "movimiento, búsqueda y acción dirigida"

        Case "Sota"
            InterpretarFigura = "inicio, aprendizaje y curiosidad"

        Case Else
            InterpretarFigura = ""
    End Select

End Function

Function InterpretarPalo(Palo As String) As String

    Select Case Palo

        Case "Bastos"
            InterpretarPalo = "energía de acción, impulso y dirección"

        Case "Copas"
            InterpretarPalo = "energía emocional, vínculos y sensibilidad"

        Case "Espadas"
            InterpretarPalo = "energía mental, análisis y decisión"

        Case "Oros"
            InterpretarPalo = "energía material, estabilidad y concreción"

        Case Else
            InterpretarPalo = ""
    End Select

End Function

Function InterpretarPlano(Plano As String) As String

    Select Case Plano

        Case "Emocional"
            InterpretarPlano = "En el plano emocional, esta energía se expresa a través de sentimientos, vínculos y sensibilidad interna"

        Case "Físico"
            InterpretarPlano = "En el plano físico, esta energía se traduce en acciones, hábitos y manifestación concreta"

        Case "Mental"
            InterpretarPlano = "En el plano mental, esta energía actúa en ideas, decisiones y procesos cognitivos"

        Case "Espiritual"
            InterpretarPlano = "En el plano espiritual, esta energía se manifiesta como propósito, visión y orientación interna"

        Case "Esencia"
            InterpretarPlano = "En la esencia, esta energía describe la vibración profunda y permanente del ser"

        Case Else
            InterpretarPlano = ""
    End Select

End Function

Function InterpretarDecanato(Decanato As Integer) As String

    Select Case Decanato

        Case 1
            InterpretarDecanato = "primer decanato: energía pura y directa del signo"

        Case 2
            InterpretarDecanato = "segundo decanato: matiz del segundo signo del elemento, aportando profundidad y complejidad"

        Case 3
            InterpretarDecanato = "tercer decanato: matiz del tercer signo del elemento, con energía de síntesis y madurez"

        Case Else
            InterpretarDecanato = ""
    End Select

End Function

Function SignoModulador(Signo As String, Decanato As Integer) As String
    Dim Elemento As String
    
    ' Determinar el elemento del signo
    Select Case UCase(Signo)
        Case "ARIES", "LEO", "SAGITARIO"
            Elemento = "FUEGO"
        Case "TAURO", "VIRGO", "CAPRICORNIO"
            Elemento = "TIERRA"
        Case "GÉMINIS", "GEMINIS", "LIBRA", "ACUARIO"
            Elemento = "AIRE"
        Case "CÁNCER", "CANCER", "ESCORPIO", "PISCIS"
            Elemento = "AGUA"
        Case Else
            SignoModulador = ""
            Exit Function
    End Select
    
    ' Determinar el signo modulador según el elemento y el decanato
    Select Case Elemento
        Case "FUEGO"
            Select Case Decanato
                Case 1: SignoModulador = "Aries"
                Case 2: SignoModulador = "Leo"
                Case 3: SignoModulador = "Sagitario"
            End Select
            
        Case "TIERRA"
            Select Case Decanato
                Case 1: SignoModulador = "Tauro"
                Case 2: SignoModulador = "Virgo"
                Case 3: SignoModulador = "Capricornio"
            End Select
            
        Case "AIRE"
            Select Case Decanato
                Case 1: SignoModulador = "Géminis"
                Case 2: SignoModulador = "Libra"
                Case 3: SignoModulador = "Acuario"
            End Select
            
        Case "AGUA"
            Select Case Decanato
                Case 1: SignoModulador = "Cáncer"
                Case 2: SignoModulador = "Escorpio"
                Case 3: SignoModulador = "Piscis"
            End Select
    End Select
End Function

