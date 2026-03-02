Attribute VB_Name = "_modASRTools3"

Option Compare Database
Option Explicit

Public DicProyecto As Object
Public DicReferencias As Object
Public DicDependencias As Object
Public DicEntorno As Object
Public ListaObjetos As Collection

Private RefActual As Object


Public Sub LeerManifiesto(ruta As String)
    Dim f As Integer
    Dim linea As String
    Dim seccion As String

    Call InicializarParser

    f = FreeFile
    Open ruta For Input As #f

'    Do While Not EOF(f)
'        Line Input #f, linea
'        linea = Trim(linea)
'
'        If linea <> "" Then
'            If Left(linea, 1) = "[" Then
'                seccion = linea
'            Else
'                Call ProcesarLinea(seccion, linea)
'            End If
'        End If
'    Loop

    Do While Not EOF(f)
        Line Input #f, linea
        linea = Trim(linea)
    
        If linea <> "" Then
    
            ' Detectar inicio de una nueva sección
            If Left(linea, 1) = "[" Then
    
                ' Si estábamos dentro de una referencia, guardarla
                If seccion = "[Referencia]" Then
                    Call GuardarReferencia
                End If
    
                ' Detectar inicio de bloque de referencia
                If linea = "[Referencia]" Then
                    Set RefActual = CreateObject("Scripting.Dictionary")
                    seccion = "[Referencia]"
                    GoTo Siguiente
                End If
    
                ' Cualquier otra sección
                seccion = linea
                GoTo Siguiente
            End If
    
            ' Procesar línea según la sección activa
            Call ProcesarLinea(seccion, linea)
    
        End If

Siguiente:
    Loop


    Close #f
End Sub


Public Sub InicializarParser()
    Set DicProyecto = CreateObject("Scripting.Dictionary")
    Set DicReferencias = CreateObject("Scripting.Dictionary")
    Set DicDependencias = CreateObject("Scripting.Dictionary")
    Set DicEntorno = CreateObject("Scripting.Dictionary")
    Set ListaObjetos = New Collection
End Sub

Private Sub ProcesarLinea(seccion As String, linea As String)
    Dim clave As String, valor As String
    Dim pos As Long

    pos = InStr(linea, "=")
    If pos = 0 Then Exit Sub

    clave = Trim(Left(linea, pos - 1))
    valor = Trim(Mid(linea, pos + 1))

    ' Quitar comillas si las hay
    If Left(valor, 1) = """" Then
        valor = Mid(valor, 2, Len(valor) - 2)
    End If

    Select Case seccion

        Case "[Proyecto]"
            DicProyecto(clave) = valor

        Case "[Referencia]"
            Call ProcesarLineaReferencia(linea)

        Case "[Dependencias]", "[Dependencias Internas]"
            Call ProcesarDependencia(linea)

        Case "[Entorno]", "[Entorno Sistema]", "[Entorno Office]", "[Idiomas]", "[Hardware]"
            DicEntorno(clave) = valor

        Case "[Objetos]"
            ListaObjetos.Add valor

    End Select
End Sub

Private Sub ProcesarLineaReferencia(linea As String)
    Dim clave As String, valor As String
    Dim pos As Long

    pos = InStr(linea, "=")
    If pos = 0 Then Exit Sub

    clave = Trim(Left(linea, pos - 1))
    valor = Trim(Mid(linea, pos + 1))

    If Left(valor, 1) = """" Then
        valor = Mid(valor, 2, Len(valor) - 2)
    End If

    RefActual(clave) = valor
End Sub

Private Sub GuardarReferencia()
    If RefActual Is Nothing Then Exit Sub
    If Not RefActual.Exists("GUID") Then Exit Sub

    Dim guid As String
    guid = Trim(RefActual("GUID"))
    If guid = "" Then Exit Sub

    ' Evitar error 450 si la clave ya existe
    If DicReferencias.Exists(guid) Then
        DicReferencias.Remove guid
    End If

    DicReferencias.Add guid, RefActual
End Sub


Private Sub ProcesarDependencia(linea As String)
    Dim obj As String, deps As String
    Dim lista As Variant
    Dim pos As Long
    Dim i As Long

    pos = InStr(linea, "=")
    obj = Trim(Left(linea, pos - 1))
    deps = Trim(Mid(linea, pos + 1))

    lista = Split(deps, ",")

    For i = LBound(lista) To UBound(lista)
        lista(i) = Trim(lista(i))
    Next

    DicDependencias(obj) = lista
End Sub

Sub ProcesarObjeto(linea As String)
    Dim tipo As String, nombre As String
    Dim pos As Long

    pos = InStr(linea, "=")
    tipo = Trim(Left(linea, pos - 1))
    nombre = Trim(Mid(linea, pos + 1))

    nombre = Replace(nombre, """", "")

    ListaObjetos.Add tipo & "|" & nombre
End Sub


'Private Sub ProcesarReferencia(linea As String)
'    Dim guid As String, ruta As String, version As String
'    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
'
'    ' GUID
'    p1 = InStr(linea, "{")
'    p2 = InStr(linea, "}")
'    guid = Mid(linea, p1, p2 - p1 + 1)
'
'    ' Ruta
'    p3 = InStr(linea, """")
'    p4 = InStr(p3 + 1, linea, """")
'    ruta = Mid(linea, p3 + 1, p4 - p3 - 1)
'
'    ' Versión
'    p1 = InStrRev(linea, "(")
'    p2 = InStrRev(linea, ")")
'    version = Mid(linea, p1 + 1, p2 - p1 - 1)
'
'    DicReferencias(guid) = ruta & "|" & version
'End Sub

