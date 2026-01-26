Attribute VB_Name = "modMsgBoxHookEx"
' ------------------------------------------------------
' Nombre:    modMsgBoxHookEx
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

' ============================================================
' API WH_CBT Hook
' ============================================================

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr

Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As LongPtr) As Long

Private Declare PtrSafe Function CallNextHookEx Lib "user32" _
        (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function GetDlgItem Lib "user32" _
        (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr

Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
        (ByVal hWnd As LongPtr, ByVal lpString As String) As Long

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As LongPtr, ByVal wMsg As Long, _
         ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

' ============================================================
' APIs HICON
' ============================================================

Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As LongPtr, _
    ByVal lpszName As String, _
    ByVal uType As Long, _
    ByVal cxDesired As Long, _
    ByVal cyDesired As Long, _
    ByVal fuLoad As Long) As LongPtr

Private Declare PtrSafe Function DestroyIcon Lib "user32" ( _
    ByVal hIcon As LongPtr) As Long

' ============================================================
' Constantes APIs
' ============================================================

Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const BM_SETIMAGE = &HF7&
Private Const IMAGE_ICON = 1

'Private Const IMAGE_ICON = 1
Private Const LR_LOADFROMFILE = &H10&
Private Const LR_DEFAULTSIZE = &H40&



' ============================================================
' Variables internas
' ============================================================


Private hHook As LongPtr
Private BtnText(1 To 7) As String
Private BtnIcon(1 To 7) As LongPtr

' ============================================================
' Resolver icono (opcional)
' ============================================================

Private Function ResolveIcon(Icon As String, TempName As String) As LongPtr
    Dim Ruta As String

    If Len(Icon) = 0 Then Exit Function

    ' 1. Obtener la ruta del .ico (tu módulo blindado)
    Ruta = ResolveIconFile_Blindado(Icon, TempName)
    If Len(Ruta) = 0 Then Exit Function

    ' 2. Convertir la ruta en un handle HICON
    ResolveIcon = LoadIconHandle(Ruta)

End Function

'Private Function ResolveIcon(Icon As String, TempName As String) As LongPtr
'    If Len(Icon) = 0 Then Exit Function
'    ResolveIcon = ResolveIconFile_Blindado(Icon, TempName)
'End Function


Private Function LoadIconHandle(RutaIco As String) As LongPtr
    If Len(Dir$(RutaIco)) = 0 Then Exit Function

    LoadIconHandle = LoadImage(0, RutaIco, IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
End Function

' ============================================================
' FUNCIÓN PRINCIPAL (estable)
' ============================================================

Public Function MsgBoxHookEx(Prompt As String, _
                             Buttons As VbMsgBoxStyle, _
                             Title As String, _
                             Optional BtOk As String = "", _
                             Optional BtCancel As String = "", _
                             Optional BtAbort As String = "", _
                             Optional BtRetry As String = "", _
                             Optional BtIgnore As String = "", _
                             Optional BtYes As String = "", _
                             Optional BtNo As String = "", _
                             Optional IconOk As String = "", _
                             Optional IconCancel As String = "", _
                             Optional IconAbort As String = "", _
                             Optional IconRetry As String = "", _
                             Optional IconIgnore As String = "", _
                             Optional IconYes As String = "", _
                             Optional IconNo As String = "") As VbMsgBoxResult
                             
Dim i As Long

    ' Textos
    BtnText(1) = BtOk
    BtnText(2) = BtCancel
    BtnText(3) = BtAbort
    BtnText(4) = BtRetry
    BtnText(5) = BtIgnore
    BtnText(6) = BtYes
    BtnText(7) = BtNo

    ' Iconos por botón (opcionales)
    BtnIcon(1) = ResolveIcon(IconOk, "msg_ok")
    BtnIcon(2) = ResolveIcon(IconCancel, "msg_cancel")
    BtnIcon(3) = ResolveIcon(IconAbort, "msg_abort")
    BtnIcon(4) = ResolveIcon(IconRetry, "msg_retry")
    BtnIcon(5) = ResolveIcon(IconIgnore, "msg_ignore")
    BtnIcon(6) = ResolveIcon(IconYes, "msg_yes")
    BtnIcon(7) = ResolveIcon(IconNo, "msg_no")

    ' Instalar hook
    hHook = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, 0, GetCurrentThreadId)

    ' Mostrar MsgBox estándar de Access
    MsgBoxHookEx = MsgBox(Prompt, Buttons, Title)

    ' Desinstalar hook
    UnhookWindowsHookEx hHook
    
    ' Liberar iconos cargados
    
    For i = 1 To 7
        If BtnIcon(i) <> 0 Then
            DestroyIcon BtnIcon(i)
            BtnIcon(i) = 0
        End If
    Next
    
End Function

' ============================================================
' CALLBACK (solo botones, estable)
' ============================================================

Private Function CBTProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If nCode = HCBT_ACTIVATE Then

        Dim s As String * 255
        GetClassName wParam, s, 255

        If Left$(s, 6) = "#32770" Then

            Dim i As Long, hBtn As LongPtr

            ' Cambiar textos e iconos de botones
            For i = 1 To 7
                hBtn = GetDlgItem(wParam, i)
                If hBtn <> 0 Then

                    If BtnText(i) <> "" Then
                        SetWindowText hBtn, BtnText(i)
                    End If

                    If BtnIcon(i) <> 0 Then
                        SendMessage hBtn, BM_SETIMAGE, IMAGE_ICON, BtnIcon(i)
                    End If

                End If
            Next

            UnhookWindowsHookEx hHook
        End If
    End If

    CBTProc = CallNextHookEx(hHook, nCode, wParam, lParam)
End Function


