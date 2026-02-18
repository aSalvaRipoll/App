Attribute VB_Name = "modBBDD_Resultados"

Option Compare Database
Option Explicit

Public Sub PersistirResultados(ByRef r As clsResultado, ByRef Incl As clsInclusion, ByRef c As clsCiclos, ByRef pd As clsPinaDes, ByRef colTr As Collection)

Dim objTr As clsTransito
Dim idTransito As Long
Dim idRes As Long

    If r.IDPersona = 0 Or r.IDFonetica = 0 Then
        Err.Raise vbObjectError + 1001, "PersistirResultados", "Resultado sin claves primarias"
    End If

    idRes = AutoNext("IDResultado", "tbuResultados", _
                             "idPersona = " & r.IDPersona & " AND " & _
                             "idFonetica = " & r.IDFonetica)

    r.IDResultado = idRes
    
    With Incl
        .IDResultado = idRes
        .IDPersona = r.IDPersona
        .IDFonetica = r.IDFonetica
        .IDInclusion = AutoNext("IDInclusion", "tbuInclusiones", _
                                "IDResultado = " & idRes & " AND " & _
                                "idPersona = " & r.IDPersona & " AND " & _
                                "idFonetica = " & r.IDFonetica)
    End With
    
    With c
        .IDResultado = idRes
        .IDPersona = r.IDPersona
        .IDFonetica = r.IDFonetica
        .idCiclo = AutoNext("IDCiclo", "tbuCiclos", _
                            "IDResultado = " & idRes & " AND " & _
                            "idPersona = " & r.IDPersona & " AND " & _
                            "idFonetica = " & r.IDFonetica)

    End With
    
    With pd
        .IDResultado = idRes
        .IDPersona = r.IDPersona
        .IDFonetica = r.IDFonetica
        .IDPinaDes = AutoNext("IDPinaDes", "tbuPinaDes", _
                              "IDResultado = " & idRes & " AND " & _
                              "idPersona = " & r.IDPersona & " AND " & _
                              "idFonetica = " & r.IDFonetica)
    End With
            
            
    idTransito = AutoNext("IDTransito", "tbuTransitos", _
                          "IDResultado = " & idRes & " AND " & _
                          "idPersona = " & r.IDPersona & " AND " & _
                          "idFonetica = " & r.IDFonetica)
                         
    For Each objTr In colTr
        With objTr
            .IDResultado = idRes
            .IDPersona = r.IDPersona
            .IDFonetica = r.IDFonetica
            .idTransito = idTransito
        End With
    Next
    
GuardarResultado r
GuardarInclusion Incl
GuardarCiclos c
GuardarPinaculosDesafios pd
GuardarTransitos colTr

End Sub


Public Sub GuardarResultado(ByRef r As clsResultado)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbuResultados", dbOpenDynaset)

    ' Generar ID manual
'    r.IDResultado = AutoNext("IDResultado", "tbuResultados", _
                                "idPersona = " & Persona.ID_Persona & _
                                " AND idFonetica = " & Fonetica.IDFonetica)

    rs.AddNew
    rs!IDResultado = r.IDResultado
    rs!IDPersona = r.IDPersona
    rs!IDFonetica = r.IDFonetica
    rs!FechaCalculo = Now

    rs!SistemaCalculo = r.SistemaCalculo
    rs!NumCiclos = r.NumCiclos
    rs!SistemaCiclos = r.SistemaCiclos
    rs!SistemaTarot = r.SistemaTarot
    rs!VersionMotor = r.VersionMotor

    rs!NumeroDestino = r.NumeroDestino
    rs!NumeroAlma = r.NumeroAlma
    rs!NumeroPersonalidad = r.NumeroPersonalidad
    rs!NumeroCaminoVida = r.NumeroCaminoVida
    rs!NumeroMadurez = r.NumeroMadurez

    rs!AnioPersonal = r.AnioPersonal
    'rs!MesPersonal = r.MesPersonal
    'rs!DiaPersonal = r.DiaPersonal
    rs!EdadPersonal = r.EdadPersonal

'    rs!CicloActual = r.CicloActual
'    rs!PinaculoActual = r.PinaculoActual
'    rs!DesafioActual = r.DesafioActual

    rs!PlanoFisico = r.PlanoFisico
    rs!PlanoEmocional = r.PlanoEmocional
    rs!PlanoMental = r.PlanoMental
    rs!PlanoIntuitivo = r.PlanoIntuitivo

    rs!PiedraAngular = r.PiedraAngular
    rs!PiedraToque = r.PiedraToque

    rs!PrimeraLetra = r.PrimeraLetra
    rs!PrimeraVocal = r.PrimeraVocal
    rs!PrimeraConsonante = r.PrimeraConsonante

    rs!RespuestaSubconsciente = r.RespuestaSubconsciente
    rs!Poder = r.NumeroPoder
'    rs!DeudaKarmica = r.DeudaKarmica

    rs.Update
    rs.Close

'    GuardarResultado = newID
    Exit Sub

ErrHandler:
    MsgBox "Error al guardar resultados: " & Err.Description, vbExclamation
'    GuardarResultado = 0
End Sub

Public Sub GuardarInclusion(ByRef inc As clsInclusion)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbuInclusiones", dbOpenDynaset)

    ' Generar ID manual
'    inc.IDInclusion = AutoNext("IDInclusion", "tbuInclusiones", _
                     "IDResultado = " & inc.IDResultado)

    
    rs.AddNew
    rs!IDInclusion = inc.IDInclusion
    rs!IDResultado = inc.IDResultado
    rs!IDPersona = inc.IDPersona
    rs!IDFonetica = inc.IDFonetica

    rs!N1 = inc.N1
    rs!N2 = inc.N2
    rs!N3 = inc.N3
    rs!N4 = inc.N4
    rs!N5 = inc.N5
    rs!N6 = inc.N6
    rs!N7 = inc.N7
    rs!N8 = inc.N8
    rs!N9 = inc.N9

    rs.Update
    rs.Close
    Exit Sub

ErrHandler:
    MsgBox "Error al guardar inclusión: " & Err.Description, vbExclamation
End Sub

Public Sub GuardarCiclos(ByRef c As clsCiclos)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
'    Dim newID As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbuCiclos", dbOpenDynaset)

    ' Generar ID manual
'    c.idCiclo = AutoNext("IDCiclo", "tbuCiclos", _
                     "IDResultado = " & c.IDResultado)

'    c.idCiclo = newID

    rs.AddNew
    rs!idCiclo = c.idCiclo
    rs!IDResultado = c.IDResultado
    rs!IDPersona = c.IDPersona

    rs!NumCiclos = c.NumCiclos
    rs!MetodoCiclos = c.MetodoCiclos

    rs!Ciclo1 = c.Ciclo1
    rs!EdadIni1 = c.EdadIni1
    rs!EdadFin1 = c.EdadFin1

    rs!Ciclo2 = c.Ciclo2
    rs!EdadIni2 = c.EdadIni2
    rs!EdadFin2 = c.EdadFin2

    rs!Ciclo3 = c.Ciclo3
    rs!EdadIni3 = c.EdadIni3
    rs!EdadFin3 = c.EdadFin3

    rs!Ciclo4 = c.Ciclo4
    rs!EdadIni4 = c.EdadIni4
    rs!EdadFin4 = c.EdadFin4

    rs.Update
    rs.Close
    Exit Sub

ErrHandler:
    MsgBox "Error al guardar ciclos: " & Err.Description, vbExclamation
End Sub

Public Sub GuardarPinaculosDesafios(ByRef p As clsPinaDes)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
'    Dim newID As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbuPinaDes", dbOpenDynaset)

    ' Generar ID manual
'    newID = AutoNext("IDPinaDes", "tbuPinaDes", _
                     "IDResultado = " & p.IDResultado)

'    p.IDPinaDes = newID

    rs.AddNew
    rs!IDPinaDes = p.IDPinaDes
    rs!IDResultado = p.IDResultado
    rs!IDPersona = p.IDPersona

    ' Pináculos
    rs!Pina1 = p.Pina1
    rs!Pina2 = p.Pina2
    rs!Pina3 = p.Pina3
    rs!Pina4 = p.Pina4

    ' Desafíos
    rs!Desa1 = p.Desa1
    rs!Desa2 = p.Desa2
    rs!Desa3 = p.Desa3
    rs!Desa4 = p.Desa4

    ' Fechas de inicio
    rs!fIni1 = p.EdadIni1
    rs!fIni2 = p.EdadIni2
    rs!fIni3 = p.EdadIni3
    rs!fIni4 = p.EdadIni4

    ' Fechas de fin
    rs!fFin1 = p.EdadFin1
    rs!fFin2 = p.EdadFin2
    rs!fFin3 = p.EdadFin3
    rs!fFin4 = p.EdadFin4

    rs.Update
    rs.Close
    Exit Sub

ErrHandler:
    MsgBox "Error al guardar Pináculos y Desafíos: " & Err.Description, vbExclamation
End Sub


Public Sub GuardarTransitos(ByRef colTransitos As Collection) ', ByVal IDResultado As Long, ByVal IDPersona As Long, ByVal IDFonetica As Long)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tr As clsTransito
    Dim newID As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbuTransitos", dbOpenDynaset)

'    tr.idTransito = AutoNext("IDTransito", "tbuTransitos", _
                         "IDResultado = " & IDResultado)
    
    ' Recorrer todos los tránsitos
    For Each tr In colTransitos

        ' Generar ID manual

        
'        tr.IDResultado = IDResultado
'        tr.IDPersona = IDPersona
'        tr.IDFonetica = IDFonetica

        rs.AddNew
        rs!idTransito = tr.idTransito
        rs!IDResultado = tr.IDResultado
        rs!IDPersona = tr.IDPersona
        rs!IDFonetica = tr.IDFonetica

        rs!orden = tr.orden
        rs!Edad = tr.Edad
        rs!anio = tr.anio

        rs!Fisico = tr.Fisico
        rs!Mental = tr.Mental
        rs!Emocional = tr.Emocional
        rs!Espiritual = tr.Espiritual
        rs!Esencia = tr.Esencia
        rs!AnioPersonal = tr.AnioPersonal

        rs!MetodoFonetico = tr.MetodoFonetico
        rs!ModoCalculo = tr.ModoCalculo

        rs!EsActual = tr.EsActual

        rs.Update
    Next tr

    rs.Close
    Exit Sub

ErrHandler:
    MsgBox "Error al guardar tránsitos: " & Err.Description, vbExclamation
End Sub


