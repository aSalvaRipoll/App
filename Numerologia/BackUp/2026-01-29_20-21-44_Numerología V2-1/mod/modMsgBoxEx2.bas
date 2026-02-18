Attribute VB_Name = "modMsgBoxEx2"
' ------------------------------------------------------
' Nombre:    modMsgBoxEx2
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

Private Const MagicWizHookKey As Long = 51488399   ' Clave interna necesaria

'Ejemplo 1
Sub EjemploClasico()

MsgBoxEx2 "¿Eliminar registro?", vbYesNo + vbQuestion, "Confirmación", _
             YesText:="&Eliminar", NoText:="&Cancelar", _
             YesIcon:="ImageMso:AcceptInvitation", _
             NoIcon:="ImageMso:Cancel"
             
End Sub
'Ejemplo2



' ============================================================
' INTERFAZ UNIFICADA: MsgBoxEx2
' ============================================================

Public Function MsgBoxEx2( _
        Prompt As String, _
        Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
        Optional Title As String = "", _
        Optional OkText As String = "", _
        Optional CancelText As String = "", _
        Optional AbortText As String = "", _
        Optional RetryText As String = "", _
        Optional IgnoreText As String = "", _
        Optional YesText As String = "", _
        Optional NoText As String = "", _
        Optional OkIcon As String = "", _
        Optional CancelIcon As String = "", _
        Optional AbortIcon As String = "", _
        Optional RetryIcon As String = "", _
        Optional IgnoreIcon As String = "", _
        Optional YesIcon As String = "", _
        Optional NoIcon As String = "" _
    ) As VbMsgBoxResult

    MsgBoxEx2 = MsgBoxHookEx( _
        Prompt, Buttons, Title, _
        BtOk:=OkText, BtCancel:=CancelText, BtAbort:=AbortText, _
        BtRetry:=RetryText, BtIgnore:=IgnoreText, _
        BtYes:=YesText, BtNo:=NoText, _
        IconOk:=OkIcon, IconCancel:=CancelIcon, IconAbort:=AbortIcon, _
        IconRetry:=RetryIcon, IconIgnore:=IgnoreIcon, _
        IconYes:=YesIcon, IconNo:=NoIcon _
    )

End Function


Public Function MesgBox(ByVal msgText As String, _
    Optional ByVal TimeInSeconds As Integer, _
    Optional ByVal intButtons = vbDefaultButton1, _
    Optional TitleText As String = "WScript") As Integer

On Error GoTo MesgBox_Err
Dim winShell As Object

Set winShell = CreateObject("WScript.Shell")

MesgBox = winShell.PopUp(msgText, TimeInSeconds, TitleText, intButtons)

MesgBox_Exit:
Exit Function

MesgBox_Err:
winShell.PopUp Err & " : " & Err.Description, 0, "MesgBox()", vbCritical
Resume MesgBox_Exit
End Function


' ============================================================
'  Módulo WizMsgBox — Cuadros de mensaje con formato
'  - Primera línea en negrita y roja
'  - Líneas siguientes normales
'  - Compatible con vbOKOnly, vbYesNo, vbOKCancel, etc.
'  - Sin hooks, sin API, sin riesgo de cierre
' ============================================================


' ------------------------------------------------------------
' Función principal
' ------------------------------------------------------------
Public Function WizMsg( _
        Title As String, _
        BoldLine1 As String, _
        Optional Line2 As String = " ", _
        Optional Line3 As String = "", _
        Optional MsgBoxStyle As VbMsgBoxStyle = vbOKOnly _
    ) As VbMsgBoxResult

    ' Activar WizHook
    WizHook.key = MagicWizHookKey

    ' Construir mensaje con separadores "@"
    Dim FullText As String
    FullText = BoldLine1 & "@" & Line2 & "@" & Line3

    ' Llamada al motor interno de Access
    WizMsg = WizHook.WizMsgBox(FullText, Title, MsgBoxStyle, 0, "")

End Function


