Attribute VB_Name = "_modASRTools2"

Option Compare Database
Option Explicit

Public DicDependencias As Object    ' Scripting.Dictionary
Public DicObjetos As Object         ' Scripting.Dictionary

Dim DicTablas As Object     ' Scripting.Dictionary
Dim DicQrys As Object       ' Scripting.Dictionary
Dim DicForms As Object      ' Scripting.Dictionary
Dim DicReports As Object    ' Scripting.Dictionary
Dim DicMacros As Object     ' Scripting.Dictionary
Dim DicModulos As Object    ' Scripting.Dictionary
Dim DicClases As Object     ' Scripting.Dictionary
Dim DicUserForms As Object  ' Scripting.Dictionary
Dim DicDesigners As Object  ' Scripting.Dictionary

Dim arrTablas, arrConsultas, arrFormularios, arrReports, arrMacros, arrDesigners, arrUserForms, arrModulos, arrClases

Private Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32" () As Long


Public Sub InicializarDependencias()
    Set DicDependencias = CreateObject("Scripting.Dictionary")
    Set DicObjetos = CreateObject("Scripting.Dictionary")
    
    Set DicTablas = CreateObject("Scripting.Dictionary")
    Set DicQrys = CreateObject("Scripting.Dictionary")
    Set DicForms = CreateObject("Scripting.Dictionary")
    Set DicReports = CreateObject("Scripting.Dictionary")
    Set DicMacros = CreateObject("Scripting.Dictionary")
    Set DicModulos = CreateObject("Scripting.Dictionary")
    Set DicClases = CreateObject("Scripting.Dictionary")
    Set DicUserForms = CreateObject("Scripting.Dictionary")
    Set DicDesigners = CreateObject("Scripting.Dictionary")
    
End Sub

Public Sub RegistrarDependencias(NombreLogico As String, RutaCarpeta As String)

    Dim fso As Object
    Dim ts As Object
    Dim texto As String
    Dim RutaArchivo As String
    Dim deps As String

    RutaArchivo = RutaCarpeta & "\" & NombreLogico

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(RutaArchivo) = False Then
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(RutaArchivo, 1, False) ' ForReading

    texto = ts.ReadAll
    ts.Close

    deps = DetectarDependencias(texto)

    If deps <> "" Then
        DicDependencias(NombreLogico) = deps
    End If

End Sub


Private Function DetectarDependencias(texto As String) As String

    Dim comp As Object 'VBIDE.VBComponent
    Dim lista As String

    For Each comp In VBE.ActiveVBProject.VBComponents
        DoEvents
        If InStr(1, texto, comp.Name, vbTextCompare) > 0 Then
            If lista <> "" Then lista = lista & ", "
            lista = lista & comp.Name
        End If
    Next comp

    DetectarDependencias = lista

End Function


Public Sub RegistrarObjeto(nombre As String, Tipo As String)

    Select Case LCase$(Tipo)
        Case "Tabla"
            DicTablas(nombre) = Tipo
        Case "Consulta"
            DicQrys(nombre) = Tipo
        Case "Formulario"
            DicForms(nombre) = Tipo
        Case "Informe"
            DicReports(nombre) = Tipo
        Case "Macro"
            DicMacros(nombre) = Tipo
        Case "módulo"
            DicModulos(nombre) = Tipo
        Case "clase"
            DicClases(nombre) = Tipo
        Case "user form"
            DicUserForms(nombre) = Tipo
        Case "designer"
            DicDesigners(nombre) = Tipo
        Case Else
            DicObjetos(nombre) = Tipo
    End Select
End Sub

Public Sub ExportarReferencias(RutaBase As String)

    Dim RutaRef As String
    'Dim f As Integer
    Dim ref As Reference
    Dim estado As String
    Dim clave As Variant
    Dim lcid As Long
    Const msoLanguageIDUI As Long = 1

    Dim vMSO As String
    Dim lIdiomaInst As Long

    Dim vSO As String
    Dim arqSO As String

    Dim cpu As String, vel As String, fab As String, modl As String
    Dim ram As String, arq As String

    Dim fso As Object 'Scripting.FileSystemObject
    Dim ts As Object 'TextStream


    Call ObtenerDatosWindows(vSO, arqSO)
    Call ObtenerHardware(cpu, vel, fab, modl, ram, arq)

    Set fso = CreateObject("Scripting.FileSystemObject")
    'Set fso = New Scripting.FileSystemObject
    
    RutaRef = RutaBase & "\proyecto.ref"
    Set ts = fso.CreateTextFile(RutaRef, True, False) ' False = ASCII, True = Unicode
    
    With ts
            
        ' --- [Proyecto] ---
        .WriteLine "[Proyecto]"
        .WriteLine "Nombre = """ & fso.GetBaseName(CurrentDb.Name) & """"
        .WriteLine "Version = " & GetProperty("Major", 0) & "." & GetProperty("Minor", 0) & "." & GetProperty("Revision", 0) & "." & GetProperty("Build", 0)
        .WriteLine "Nombre VBA = """ & VBE.ActiveVBProject.Name & """"
        
        .WriteLine "VersionAccess = """ & Application.version & """"
        .WriteLine "OptionCompare = ""Database"""
        .WriteLine "OptionExplicit = ""True"""
        .WriteLine "FechaExportacion = """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """"
        .WriteLine ""

        ' --- [Hardware] ---
        .WriteLine "[Hardware]"
        .WriteLine "CPU = """ & cpu & """"
        .WriteLine "Velocidad = """ & vel & """"
        .WriteLine "Fabricante = """ & fab & """"
        .WriteLine "Modelo = """ & modl & """"
        .WriteLine "RAM = """ & ram & """"
        .WriteLine "Arquitectura = """ & arq & """"

        ' --- [Entorno] ---
        .WriteLine "[Entorno Sistema]"
        .WriteLine "SO = """ & vSO & """"
        .WriteLine "Arquitectura = """ & arqSO & """"
        .WriteLine ""
    
        .WriteLine "[Entorno Office]"
        .WriteLine "Version Access = """ & Application.version & " (" & Application.Build & ")"""
'        .WriteLine "Version MSO = """ & ObtenerVersionMSO_WMI & """"

#If Win64 Then
        .WriteLine "Arquitectura Office = ""64 bits"""
#Else
        .WriteLine "Arquitectura Office = ""32 bits"""
#End If
        .WriteLine "Version ACE = """ & CurrentDb.version & """"
        .WriteLine "Version ACEDAO = """ & DAO.DBEngine.version & """"
        .WriteLine "Version VBE = """ & VBE.version & """"
        .WriteLine ""
    
        ' --- [Idiomas] ---
        .WriteLine "[Idiomas]"
        .WriteLine "Idioma Windows = """ & LCIDaCodigo(GetUserDefaultLCID()) & """"
        .WriteLine "Idioma Acccess = """ & LCIDaCodigo(Application.LanguageSettings.LanguageID(msoLanguageIDUI)) & """"
        .WriteLine ""

        ' --- [Referencia] ---
        For Each ref In Application.References
            DoEvents
            If ref.IsBroken Then
                estado = "MISSING"
            Else
                estado = "OK"
            End If

            .WriteLine "[Referencia]"
            .WriteLine "Nombre = """ & ref.Name & """"
            .WriteLine "GUID = """ & ref.guid & """"
            .WriteLine "Version = """ & ref.Major & "." & ref.Minor & """"
            On Error Resume Next
            .WriteLine "Ruta = """ & ref.FullPath & """"
            On Error GoTo 0
            .WriteLine "Estado = """ & estado & """"
            .WriteLine ""
        Next ref
        
        ' --- [Objetos] ---
        Call AñadirGrupo(DicObjetos, DicTablas)
        Call AñadirGrupo(DicObjetos, DicQrys)
        Call AñadirGrupo(DicObjetos, DicForms)
        Call AñadirGrupo(DicObjetos, DicReports)
        Call AñadirGrupo(DicObjetos, DicMacros)
        Call AñadirGrupo(DicObjetos, DicDesigners)
        Call AñadirGrupo(DicObjetos, DicUserForms)
        Call AñadirGrupo(DicObjetos, DicModulos)
        Call AñadirGrupo(DicObjetos, DicClases)

        Set DicTablas = Nothing
        Set DicQrys = Nothing
        Set DicForms = Nothing
        Set DicReports = Nothing
        Set DicMacros = Nothing
        Set DicModulos = Nothing
        Set DicClases = Nothing
        Set DicUserForms = Nothing
        Set DicDesigners = Nothing

        estado = ""
        .WriteLine "[Objetos]"
        For Each clave In DicObjetos.Keys
            DoEvents
            If (DicObjetos(clave) <> estado) And (estado <> "") Then
                .WriteLine ""
            End If
            '.WriteLine clave & " = """ & DicObjetos(clave) & """"
            .WriteLine DicObjetos(clave) & " = " & clave
            If (DicObjetos(clave) <> estado) Then _
                estado = DicObjetos(clave)
        Next clave
        .WriteLine ""

        ' --- [Dependencias] ---
        estado = ""
        .WriteLine "[Dependencias]"
        
        For Each clave In DicObjetos.Keys
            DoEvents
            If DicDependencias.Exists(clave) Then
                If (DicObjetos(clave) <> estado) And (estado <> "") Then
                    .WriteLine ""
                End If
                .WriteLine DicObjetos(clave) & " | " & clave & " = """ & DicDependencias(clave) & """"
                If (DicObjetos(clave) <> estado) Then _
                    estado = DicObjetos(clave)
            End If
        Next clave
        .WriteLine ""

        .Close
    End With
    
End Sub

Private Sub AñadirGrupo(dicDestino As Object, dicOrigen As Object)
    Dim claves As Variant
    Dim i As Long
    Dim clave As Variant

    ' Obtener claves ordenadas
    claves = dicOrigen.Keys
    Call OrdenarArray(claves)

    ' Añadir al diccionario maestro en orden
    For Each clave In claves
        dicDestino(clave) = dicOrigen(clave)
    Next clave
    
    dicOrigen.RemoveAll
    
End Sub

Private Sub OrdenarArray(arr As Variant)
    Dim i As Long, j As Long
    Dim Temp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                Temp = arr(i)
                arr(i) = arr(j)
                arr(j) = Temp
            End If
        Next j
    Next i
End Sub





Function LCIDaCodigo(lcid As Long) As String
    Select Case lcid
        Case 3082: LCIDaCodigo = "es-ES"          ' Español internacional
        Case 1034: LCIDaCodigo = "es-ES-trad"     ' Español tradicional
        Case 1027: LCIDaCodigo = "ca-ES"          ' Catalán
        Case 2051: LCIDaCodigo = "ca-ES-valencia" ' Valenciano
        Case 1069: LCIDaCodigo = "eu-ES"          ' Euskera
        Case 1110: LCIDaCodigo = "gl-ES"          ' Gallego
        Case 1033: LCIDaCodigo = "en-US"
        Case 1036: LCIDaCodigo = "fr-FR"

        Case Else: LCIDaCodigo = "desconocido"
    End Select
End Function

Sub ObtenerHardware(ByRef cpu As String, ByRef Velocidad As String, _
                    ByRef Fabricante As String, ByRef Modelo As String, _
                    ByRef ram As String, ByRef Arquitectura As String)

    On Error GoTo ErrHandler

    Dim wmi As Object
    Dim col As Object
    Dim obj As Object

    Set wmi = GetObject("winmgmts:\\.\root\cimv2")

    ' CPU
    Set col = wmi.ExecQuery("Select * from Win32_Processor")
    For Each obj In col
        cpu = obj.Name
        Velocidad = CStr(obj.MaxClockSpeed) & " MHz"
        Exit For
    Next

    ' Sistema
    Set col = wmi.ExecQuery("Select * from Win32_ComputerSystem")
    For Each obj In col
        Fabricante = obj.Manufacturer
        Modelo = obj.Model
        ram = Format(obj.TotalPhysicalMemory / 1024 / 1024 / 1024, "0.00") & " GB"
        Exit For
    Next

    ' Arquitectura
    Set col = wmi.ExecQuery("Select * from Win32_OperatingSystem")
    For Each obj In col
        Arquitectura = obj.OSArchitecture
        Exit For
    Next

    Exit Sub

ErrHandler:
    cpu = "desconocido"
    Velocidad = "desconocido"
    Fabricante = "desconocido"
    Modelo = "desconocido"
    ram = "desconocido"
    Arquitectura = "desconocida"
End Sub


Function ObtenerDatosWindows(ByRef VersionSO As String, ByRef ArquitecturaSO As String)
    On Error GoTo ErrHandler

    Dim objWMI As Object
    Dim colOS As Object
    Dim objOS As Object

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colOS = objWMI.ExecQuery("Select * from Win32_OperatingSystem")

    For Each objOS In colOS
        VersionSO = objOS.Caption & " (" & objOS.version & "). Build: " & objOS.BuildNumber
        ArquitecturaSO = objOS.OSArchitecture
        Exit For
    Next

    Exit Function

ErrHandler:
    VersionSO = "desconocido"
    ArquitecturaSO = "desconocida"
End Function



'Sub ObtenerDatosExcel(ByRef VersionMSO As String, ByRef IdiomaInstalacion As Long)
'    On Error GoTo ErrHandler
'
'    Dim xl As Object
'    Set xl = CreateObject("Excel.Application")
'
'    ' Versión real del motor MSO
'    VersionMSO = xl.Application.MSOFileVersion
'
'    ' Idioma de instalación de Office
'    IdiomaInstalacion = xl.LanguageSettings.LanguageID(1) ' msoLanguageIDInstall = 1
'
'    xl.Quit
'    Set xl = Nothing
'    Exit Sub
'
'ErrHandler:
'    VersionMSO = "desconocido"
'    IdiomaInstalacion = 0
'End Sub


'Function ObtenerVersionMSO() As String
'    On Error GoTo ErrHandler
'
'    Dim xl As Object
'    Set xl = CreateObject("Excel.Application")
'
'    ObtenerVersionMSO = xl.Application.MSOFileVersion
'
'    xl.Quit
'    Set xl = Nothing
'    Exit Function
'
'ErrHandler:
'    ObtenerVersionMSO = "desconocido"
'End Function
'
'Function ObtenerIdiomaInstalacionOffice() As Long
'    On Error GoTo ErrHandler
'
'    Dim xl As Object
'    Set xl = CreateObject("Excel.Application")
'
'    ObtenerIdiomaInstalacionOffice = xl.LanguageSettings.LanguageID(1) ' msoLanguageIDInstall = 1
'
'    xl.Quit
'    Set xl = Nothing
'    Exit Function
'
'ErrHandler:
'    ObtenerIdiomaInstalacionOffice = 0 ' desconocido
'End Function

'Public Sub ExportarReferencias(RutaBase As String)
'
'    Dim RutaRef As String
'    'Dim f As Integer
'    Dim ref As Reference
'    Dim estado As String
'    Dim clave As Variant
'    Dim lcid As Long
'    Const msoLanguageIDUI As Long = 1
'
'    Dim vMSO As String
'    Dim lIdiomaInst As Long
'
'    Dim vSO As String
'    Dim arqSO As String
'
'    Dim cpu As String, vel As String, fab As String, modl As String
'    Dim ram As String, arq As String
'
'    Dim fso As Object 'Scripting.FileSystemObject
'    Dim ts As Object 'TextStream
'
'
'    Call ObtenerDatosWindows(vSO, arqSO)
'    Call ObtenerHardware(cpu, vel, fab, modl, ram, arq)
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    'Set fso = New Scripting.FileSystemObject
'
'    RutaRef = RutaBase & "\proyecto.ref"
'    Set ts = fso.CreateTextFile(RutaRef, True, False) ' False = ASCII, True = Unicode
'
''    f = FreeFile
'
''    'Call ObtenerDatosExcel(vMSO, lIdiomaInst)
'
''    Open RutaRef For Output As #f
'
'    ' --- [Proyecto] ---
''    Print #f, "[Proyecto]"
''    Print #f, "Nombre = """ & VBE.ActiveVBProject.Name & """"
''    Print #f, "VersionAccess = """ & Application.version & """"
''    Print #f, "OptionCompare = ""Database"""
''    Print #f, "OptionExplicit = ""True"""
''    Print #f, "FechaExportacion = """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """"
''    Print #f, ""
'
'    With ts
'Debug.Print
'Debug.Print "[Proyecto]"
'        .WriteLine "[Proyecto]"
'        .WriteLine "Nombre = """ & fso.GetBaseName(CurrentDb.Name) & """"
'        .WriteLine "Nombre VBA = """ & VBE.ActiveVBProject.Name & """"
'        .WriteLine "VersionAccess = """ & Application.version & """"
'        .WriteLine "OptionCompare = ""Database"""
'        .WriteLine "OptionExplicit = ""True"""
'        .WriteLine "FechaExportacion = """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """"
'        .WriteLine ""
'
'
'
''    ' --- [Hardware] ---
''    Print #f, "[Hardware]"
''    Print #f, "CPU = """ & cpu & """"
''    Print #f, "Velocidad = """ & vel & """"
''    Print #f, "Fabricante = """ & fab & """"
''    Print #f, "Modelo = """ & modl & """"
''    Print #f, "RAM = """ & ram & """"
''    Print #f, "Arquitectura = """ & arq & """"
'
'
''Debug.Print
'Debug.Print "[Hardware]"
'        ' --- [Hardware] ---
'        .WriteLine "[Hardware]"
'        .WriteLine "CPU = """ & cpu & """"
'        .WriteLine "Velocidad = """ & vel & """"
'        .WriteLine "Fabricante = """ & fab & """"
'        .WriteLine "Modelo = """ & modl & """"
'        .WriteLine "RAM = """ & ram & """"
'        .WriteLine "Arquitectura = """ & arq & """"
'
'
''    ' --- [Entorno] ---
''    Print #f, "[Entorno Sistema]"
''    Print #f, "SO = """ & vSO & """"
''    Print #f, "Arquitectura = """ & arqSO & """"
'''    Print #f, "SO = """ & Environ$("OS") & """"
'''
'''    If Environ$("PROCESSOR_ARCHITEW6432") <> "" Then
'''        Print #f, "Arquitectura = ""x64"""
'''    Else
'''        arq = Environ$("PROCESSOR_ARCHITECTURE")
'''    End If
''    'Print #f, "Arquitectura = """ & Environ$("PROCESSOR_ARCHITECTURE") & """"
''    'Print #f, "Usuario = """ & Environ$("USERNAME") & """"
'
'
''    ' --- [Entorno] ---
''    Print #f, "[Entorno Sistema]"
''    Print #f, "SO = """ & vSO & """"
''    Print #f, "Arquitectura = """ & arqSO & """"
''    Print #f, ""
''
''    Print #f, "[Entorno Office]"
''    Print #f, "Version Access = """ & Application.version & " (" & Application.Build & ")"""
''    Print #f, "Version MSO = """ & ObtenerVersionMSO_WMI & """"
''
''#If Win64 Then
''    Print #f, "Arquitectura Office = ""64 bits"""
''#Else
''    Print #f, "Arquitectura Office = ""32 bits"""
''#End If
''    Print #f, "Version ACE = """ & CurrentDb.version & """"
''    Print #f, "Version ACEDAO = """ & DAO.DBEngine.version & """"
''    Print #f, "Version VBE = """ & VBE.version & """"
''    Print #f, ""
''
''    'Print #f, "Version Office = """ & ObtenerVersionMSO & """"
''    'Print #f, "Version MSO = """ & vMSO & """"
'
''Debug.Print
'Debug.Print "[Entorno Sistema]"
'        ' --- [Entorno] ---
'        .WriteLine "[Entorno Sistema]"
'        .WriteLine "SO = """ & vSO & """"
'        .WriteLine "Arquitectura = """ & arqSO & """"
'        .WriteLine ""
'
'        .WriteLine "[Entorno Office]"
'        .WriteLine "Version Access = """ & Application.version & " (" & Application.Build & ")"""
''        .WriteLine "Version MSO = """ & ObtenerVersionMSO_WMI & """"
'
'#If Win64 Then
'        .WriteLine "Arquitectura Office = ""64 bits"""
'#Else
'        .WriteLine "Arquitectura Office = ""32 bits"""
'#End If
'        .WriteLine "Version ACE = """ & CurrentDb.version & """"
'        .WriteLine "Version ACEDAO = """ & DAO.DBEngine.version & """"
'        .WriteLine "Version VBE = """ & VBE.version & """"
'        .WriteLine ""
'
'
''    Print #f, "[Idiomas]"
''    Print #f, "Idioma Windows = """ & LCIDaCodigo(GetUserDefaultLCID()) & """"
''    Print #f, "Idioma Acccess = """ & LCIDaCodigo(Application.LanguageSettings.LanguageID(msoLanguageIDUI)) & """"
''    'Print #f, "Idioma Office = """ & LCIDaCodigo(lIdiomaInst) & """"
''
''    'Print #f, "Idioma Office = """ & LCIDaCodigo(WizHook.GetOption(126)) & """"
''    'Print #f, "Idioma Office = """ & LCIDaCodigo(WizHook.CurrentLangID) & """"
''
''    'Print #f, "LCID = """ & WizHook.GetOption(126) & """"
'
''Debug.Print
'Debug.Print "[Idiomas]"
'        ' --- [Idiomas] ---
'        .WriteLine "[Idiomas]"
'        .WriteLine "Idioma Windows = """ & LCIDaCodigo(GetUserDefaultLCID()) & """"
'        .WriteLine "Idioma Acccess = """ & LCIDaCodigo(Application.LanguageSettings.LanguageID(msoLanguageIDUI)) & """"
'        .WriteLine ""
'
''Debug.Print
'Debug.Print "[Referencias]"
'        ' --- [Referencia] ---
'        For Each ref In Application.References
'            DoEvents
'            If ref.IsBroken Then
'                estado = "MISSING"
'            Else
'                estado = "OK"
'            End If
'
''        Print #f, "[Referencia]"
''        Print #f, "Nombre = """ & ref.Name & """"
''        Print #f, "GUID = """ & ref.guid & """"
''        Print #f, "Version = """ & ref.Major & "." & ref.Minor & """"
''        On Error Resume Next
''        Print #f, "Ruta = """ & ref.FullPath & """"
''        On Error GoTo 0
''        Print #f, "Estado = """ & estado & """"
''        Print #f, ""
'
'            .WriteLine "[Referencia]"
'            .WriteLine "Nombre = """ & ref.Name & """"
'            .WriteLine "GUID = """ & ref.guid & """"
'            .WriteLine "Version = """ & ref.Major & "." & ref.Minor & """"
'            On Error Resume Next
'            .WriteLine "Ruta = """ & ref.FullPath & """"
'            On Error GoTo 0
'            .WriteLine "Estado = """ & estado & """"
'            .WriteLine ""
'        Next ref
'
''Debug.Print
'Debug.Print "[Objetos]"
'
''        arrTablas = ClavesOrdenadas(DicTablas)
''        arrQrys = ClavesOrdenadas(DicQrys)
''        arrForms = ClavesOrdenadas(DicFormu)
''        arrReports = ClavesOrdenadas(DicReports)
''        arrMacros = ClavesOrdenadas(DicMacros)
''        arrDesigners = ClavesOrdenadas(DicDesigners)
''        arrUserForms = ClavesOrdenadas(DicUserForms)
''        arrModulos = ClavesOrdenadas(DicModulos)
''        arrClases = ClavesOrdenadas(DicClases)
'
''        Call AñadirGrupo(DicObjetos, arrTablas, DicTablas)
''        Call AñadirGrupo(DicObjetos, arrQrys, DicQrys)
''        Call AñadirGrupo(DicObjetos, arrForms, DicFormu)
''        Call AñadirGrupo(DicObjetos, arrReports, DicReports)
''        Call AñadirGrupo(DicObjetos, arrMacros, DicMacros)
''        Call AñadirGrupo(DicObjetos, arrDesigners, DicDesigners)
''        Call AñadirGrupo(DicObjetos, arrUserForms, DicUserForms)
''        Call AñadirGrupo(DicObjetos, arrModulos, DicModulos)
''        Call AñadirGrupo(DicObjetos, arrClases, DicClases)
'
'        Call AñadirGrupo(DicObjetos, DicTablas)
'        Call AñadirGrupo(DicObjetos, DicQrys)
'        Call AñadirGrupo(DicObjetos, DicFormu)
'        Call AñadirGrupo(DicObjetos, DicReports)
'        Call AñadirGrupo(DicObjetos, DicMacros)
'        Call AñadirGrupo(DicObjetos, DicDesigners)
'        Call AñadirGrupo(DicObjetos, DicUserForms)
'        Call AñadirGrupo(DicObjetos, DicModulos)
'        Call AñadirGrupo(DicObjetos, DicClases)
'
'        Set DicTablas = Nothing
'        Set DicQrys = Nothing
'        Set DicForms = Nothing
'        Set DicReports = Nothing
'        Set DicMacros = Nothing
'        Set DicModulos = Nothing
'        Set DicClases = Nothing
'        Set DicUserForms = Nothing
'        Set DicDesigners = Nothing
'
''        For Each clave In DicModulos.Keys
''            DicObjetos(clave) = DicModulos(clave)
''        Next clave
''        For Each clave In DicClases.Keys
''            DicObjetos(clave) = DicClases(clave)
''        Next clave
''        For Each clave In DicUserForms.Keys
''            DicObjetos(clave) = DicUserForms(clave)
''        Next clave
''        For Each clave In DicDesigners.Keys
''            DicObjetos(clave) = DicDesigners(clave)
''        Next clave
'
'
'        ' --- [Objetos] ---
'        estado = ""
'        .WriteLine "[Objetos]"
'        For Each clave In DicObjetos.Keys
'            DoEvents
'            If (DicObjetos(clave) <> estado) And (estado <> "") Then
'                .WriteLine ""
'            End If
'            '.WriteLine clave & " = """ & DicObjetos(clave) & """"
'            .WriteLine DicObjetos(clave) & " = " & clave
'            If (DicObjetos(clave) <> estado) Then _
'                estado = DicObjetos(clave)
'        Next clave
'        .WriteLine ""
'
'
'
'    ' --- [Dependencias Internas] ---
''    Print #f, "[Dependencias]"
''    For Each clave In DicDependencias.Keys
''        Print #f, clave & " = """ & DicDependencias(clave) & """"
''    Next clave
''    Print #f, ""
'
''Debug.Print
'Debug.Print "[Dependencias]"
'        ' --- [Dependencias] ---
'        estado = ""
'        .WriteLine "[Dependencias]"
'        'For Each clave In DicDependencias.Keys
'        For Each clave In DicObjetos.Keys
'            DoEvents
'            If DicDependencias.Exists(clave) Then
'                If (DicObjetos(clave) <> estado) And (estado <> "") Then
'                    .WriteLine ""
'                End If
'                .WriteLine DicObjetos(clave) & " | " & clave & " = """ & DicDependencias(clave) & """"
'                If (DicObjetos(clave) <> estado) Then _
'                    estado = DicObjetos(clave)
'            End If
'        Next clave
'        .WriteLine ""
'
'
'
''    Close #f
'        .Close
'    End With
'
'End Sub


'Public Sub RegistrarDependencias(NombreLogico As String, RutaCarpeta As String)
'
'    Dim f As Integer
'    Dim Texto As String
'    Dim RutaArchivo As String
'    Dim deps As String
'
'    RutaArchivo = RutaCarpeta & "\" & NombreLogico
'
'    f = FreeFile
'
'
'    Open RutaArchivo For Input As #f
'
'    If LOF(f) = 0 Then
'        DicDependencias(NombreLogico) = "Archivo vacío"
'        Close #f
'        Exit Sub
'    End If
'
'
'    Texto = Input$(LOF(f), f)
'    Close #f
'
'    deps = DetectarDependencias(Texto)
'
'    If deps <> "" Then
'        DicDependencias(NombreLogico) = deps
'    End If
'
'End Sub

'Public Sub RegistrarObjeto(Tipo As String, Nombre As String)
'    ListaObjetos.Add Tipo & "|" & Nombre
'End Sub

'Private Sub AñadirGrupo(dicDestino As Object, arrClaves As Variant, dicOrigen As Object)
'    Dim clave As Variant
'    For Each clave In arrClaves
'        dicDestino(clave) = dicOrigen(clave)
'    Next clave
'End Sub


'Function ObtenerVersionMSO_WMI() As String
'    On Error GoTo ErrHandler
'
'    Dim wmi As Object
'    Dim col As Object
'    Dim obj As Object
'    Dim versionMax As String
'
'    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
'
'    ' Buscar cualquier MSO.DLL en el sistema
'    Set col = wmi.ExecQuery( _
'        "SELECT Version FROM CIM_DataFile WHERE Name LIKE '%\\MSO.DLL'" _
'    )
'
'    versionMax = ""
'
'    For Each obj In col
'        If obj.version <> "" Then
'            If versionMax = "" Then
'                versionMax = obj.version
'            Else
'                ' Comparar versiones como cadenas numéricas
'                If CDbl(Left(obj.version, InStr(obj.version, ".") - 1)) > _
'                   CDbl(Left(versionMax, InStr(versionMax, ".") - 1)) Then
'                    versionMax = obj.version
'                End If
'            End If
'        End If
'    Next
'
'    ObtenerVersionMSO_WMI = versionMax
'    Exit Function
'
'ErrHandler:
'    ObtenerVersionMSO_WMI = ""
'End Function


'Public Function ClavesOrdenadas(dic As Object) As Variant
'    Dim arr As Variant
'    Dim i As Long, j As Long
'    Dim temp As String
'
'    arr = dic.Keys
'
'    ' Burbuja (suficiente para diccionarios pequeños/medianos)
'    For i = LBound(arr) To UBound(arr) - 1
'        For j = i + 1 To UBound(arr)
'            If arr(j) < arr(i) Then
'                temp = arr(i)
'                arr(i) = arr(j)
'                arr(j) = temp
'            End If
'        Next j
'    Next i
'
'    ClavesOrdenadas = arr
'End Function
