Attribute VB_Name = "04_modFunciones"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 04_modFunciones
' Funciones generales del Inspector (Unicode, formato, texto, etc.)
'---------------------------------------------------------------
' Este módulo centraliza utilidades transversales:
'   - Iconos Unicode (desde tblUnicode)
'   - Formatos visuales para el Inspector
'   - Funciones de texto
'   - Funciones de validación
'   - Funciones auxiliares para columnas
'   - Preparado para ampliaciones futuras
'===============================================================


'===============================================================
' UNICODE / ICONOS
'===============================================================

'---------------------------------------------------------------
' Obtener un icono Unicode por nombre
'---------------------------------------------------------------
Public Function IconoUnicode(nombre As String) As String
    IconoUnicode = Nz(DLookup("Texto", "tblUnicode", "Nombre='" & nombre & "'"), "")
End Function

'---------------------------------------------------------------
' Icono + texto (formato estándar del Inspector)
'---------------------------------------------------------------
Public Function IconoTexto(nombreIcono As String, texto As String) As String
    IconoTexto = IconoUnicode(nombreIcono) & " " & texto
End Function

'---------------------------------------------------------------
' Icono para severidad
'---------------------------------------------------------------
Public Function IconoSeveridad(severidad As Long) As String
    Select Case severidad
        Case sevInfo:  IconoSeveridad = IconoTexto("Info", "INFO")
        Case sevAviso: IconoSeveridad = IconoTexto("Aviso", "AVISO")
        Case sevError: IconoSeveridad = IconoTexto("Error", "ERROR")
        Case Else:     IconoSeveridad = IconoTexto("Info", "DESCONOCIDO")
    End Select
End Function

'---------------------------------------------------------------
' Icono para tipo de elemento (módulo, clase, formulario, etc.)
'---------------------------------------------------------------
Public Function IconoElemento(tipo As String) As String
    Select Case LCase(tipo)
        Case "modulo", "module":     IconoElemento = IconoUnicode("Modulo")
        Case "clase", "class":       IconoElemento = IconoUnicode("Clase")
        Case "formulario", "form":   IconoElemento = IconoUnicode("Carpeta")
        Case "informe", "report":    IconoElemento = IconoUnicode("Archivo")
        Case Else:                   IconoElemento = IconoUnicode("Archivo")
    End Select
End Function

'---------------------------------------------------------------
' Icono para miembros (funciones, subs, propiedades)
'---------------------------------------------------------------
Public Function IconoMiembro(tipo As String) As String
    Select Case LCase(tipo)
        Case "sub", "procedure", "function", "property"
            IconoMiembro = IconoUnicode("Funcion")
        Case Else
            IconoMiembro = IconoUnicode("Funcion")
    End Select
End Function

'---------------------------------------------------------------
' Iconos para flechas
'---------------------------------------------------------------
Public Function IconoFlechaArriba() As String
    IconoFlechaArriba = IconoUnicode("FlechaArriba")
End Function

Public Function IconoFlechaAbajo() As String
    IconoFlechaAbajo = IconoUnicode("FlechaAbajo")
End Function


'===============================================================
' FORMATO VISUAL
'===============================================================

'---------------------------------------------------------------
' Formato completo para la columna "Elemento"
'---------------------------------------------------------------
Public Function FormatoElemento(nombreElemento As String, tipo As String) As String
    FormatoElemento = IconoElemento(tipo) & " " & nombreElemento
End Function

'---------------------------------------------------------------
' Formato completo para la columna "Miembro"
'---------------------------------------------------------------
Public Function FormatoMiembro(nombreMiembro As String, tipo As String) As String
    FormatoMiembro = IconoMiembro(tipo) & " " & nombreMiembro
End Function

'---------------------------------------------------------------
' Truncar texto para columnas largas
'---------------------------------------------------------------
Public Function TruncarTexto(texto As String, maxLen As Long) As String
    If Len(texto) <= maxLen Then
        TruncarTexto = texto
    Else
        TruncarTexto = Left$(texto, maxLen - 1) & "…"
    End If
End Function


'===============================================================
' TEXTO / VALIDACIÓN
'===============================================================

'---------------------------------------------------------------
' Normalizar nombres (quitar espacios, tabs, etc.)
'---------------------------------------------------------------
Public Function NormalizarNombre(texto As String) As String
    NormalizarNombre = Trim$(Replace$(texto, vbTab, ""))
End Function

'---------------------------------------------------------------
' ¿Es un nombre de miembro privado?
'---------------------------------------------------------------
Public Function EsPrivado(nombre As String) As Boolean
    EsPrivado = (LCase$(Left$(nombre, 1)) = "_")
End Function

'---------------------------------------------------------------
' ¿Es un módulo de clase?
'---------------------------------------------------------------
Public Function EsModuloClase(nombre As String) As Boolean
    EsModuloClase = (LCase$(Right$(nombre, 6)) = ".class")
End Function


'===============================================================
' LISTAS / COLECCIONES
'===============================================================

'---------------------------------------------------------------
' Convertir una colección en un array (útil para ordenaciones)
'---------------------------------------------------------------
Public Function ColeccionAArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long

    ReDim arr(1 To col.Count)

    For i = 1 To col.Count
        Set arr(i) = col(i)
    Next i

    ColeccionAArray = arr
End Function

