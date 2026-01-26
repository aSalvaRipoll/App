Attribute VB_Name = "modMsgBoxExPlus"
' ------------------------------------------------------
' Nombre:    modMsgBoxExPlus
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

Private Const MagicWizHookKey As Long = 51488399

' ============================================================
' MsgBoxExPlus — Cuadros de mensaje modernos con:
' - Primera línea en negrita y roja
' - Hasta 5 líneas
' - Iconos propios
' - Botones personalizados
' ============================================================

Public Function MsgBoxExPlus( _
        Optional Titulo As String = "", _
        Optional Linea1 As String = "", _
        Optional Linea2 As String = "", _
        Optional Linea3 As String = "", _
        Optional Linea4 As String = "", _
        Optional Linea5 As String = "", _
        Optional Estilo As VbMsgBoxStyle = 0) As VbMsgBoxResult
        'Optional Icono As Variant = vbInformation, _
        Optional Boton1 As String = "Aceptar", _
        Optional Boton2 As String = "", _
        Optional Boton3 As String = "", _
        Optional ValorBoton1 As Long = vbOK, _
        Optional ValorBoton2 As Long = vbCancel, _
        Optional ValorBoton3 As Long = vbAbort _
    ) As VbMsgBoxResult

    Dim texto As String
    Dim Botones As String
    Dim IconoInterno As Variant

    ' Activar WizHook
    WizHook.key = MagicWizHookKey

    ' Construir texto con separadores "@"
    texto = Linea1 & "@" & Linea2 & "@" & Linea3 & "@" & Linea4 & "@" & Linea5

'    ' ------------------------------------------------------------
'    ' ICONO
'    ' ------------------------------------------------------------
'    ' Si Icono es un archivo .ico ? usarlo directamente
'
'    ' ICONO
'    IconoInterno = ResolveIconMSO(Icono)
'
''    If VarType(Icono) = vbString Then
''        IconoInterno = Icono
''    Else
''        ' Si es vbInformation, vbExclamation, etc.
''        IconoInterno = CInt(Icono)
''    End If
'
'    ' ------------------------------------------------------------
'    ' BOTONES PERSONALIZADOS
'    ' ------------------------------------------------------------
'    ' Formato: "Texto1;Valor1|Texto2;Valor2|Texto3;Valor3"
'    Botones = Boton1 & ";" & ValorBoton1
'
'    If Boton2 <> "" Then
'        Botones = Botones & "|" & Boton2 & ";" & ValorBoton2
'    End If
'
'    If Boton3 <> "" Then
'        Botones = Botones & "|" & Boton3 & ";" & ValorBoton3
'    End If

    ' ------------------------------------------------------------
    ' LLAMADA FINAL
    ' ------------------------------------------------------------
    MsgBoxExPlus = WizHook.WizMsgBox(texto, Titulo, Estilo, 0, "")
End Function

Private Function ResolveIconMSO(Icono As Variant) As Variant
    If VarType(Icono) <> vbString Then
        ResolveIconMSO = Icono
        Exit Function
    End If

    If LCase$(Right$(Icono, 4)) = ".ico" Then
        ResolveIconMSO = Icono
        Exit Function
    End If

    Dim TempName As String
    TempName = "msgbox_mso"
    ResolveIconMSO = ResolveIconFile_Blindado(CStr(Icono), TempName)
End Function

'Private Function ResolveIconMSO(Icono As Variant) As Variant
'    ' Si no es texto, devolver tal cual (vbInformation, vbExclamation…)
'    Dim TempName As String
'
'    TempName = "msgbox_mso"
'    ResolveIconFile_Blindado Icono, TempName
'
'
'    If VarType(Icono) <> vbString Then
'        ResolveIconMSO = Icono
'        Exit Function
'    End If
'
'    ' Si es ruta a archivo .ico, devolverla tal cual
'    If LCase$(Right$(Icono, 4)) = ".ico" Then
'        ResolveIconMSO = Icono
'        Exit Function
'    End If
'
'    ' Si es un ImageMSO ? convertirlo a .ico temporal
'    ResolveIconMSO = ResolveIconFile_Blindado(Icono, "msgbox_mso")
'End Function

