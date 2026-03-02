Attribute VB_Name = "_basDebugMotores"
Option Compare Database
Option Explicit

Public strDebug As String

Public DebugMotor As Boolean
Public DebugDTO As Boolean

'====================================================================
'====================================================================
'====================================================================

' ============================================================
'   Rutina auxiliar de diagnóstico del motor
'   Imprime el estado completo del DTO
' ============================================================
Public Sub MF_DebugDTO(Proc As String)

'    strDebug = ""

    If DebugDTO Then
        If Not DebugMotor Then
            strDebug = ""
        End If
        
        If ObjDTO Is Nothing Then
            addLog "DTO no inicializado."
            Exit Sub
        End If
    
        addLog
        addLog "==============================="
        addLog "     ESTADO DEL DTO"
        addLog "==============================="
    
        addLog
        addLog "-------------------------------"
        addLog " Proc.: " & Proc
        addLog "-------------------------------"
        addLog
        
        addLog "Texto original:        " & ObjDTO.TextoOriginal
        addLog
        addLog "Texto Corregido:       " & ObjDTO.TextoCorregido
        addLog
        addLog "Texto normalizado:     " & ObjDTO.TextoNormalizado
        addLog
        addLog "-------------------------------"
        addLog
        addLog "SilabasAuto:           " & ObjDTO.SilabasAuto
        addLog
        addLog "SilabasAcentuadas:     " & ObjDTO.SilabasAcentuadas
        addLog
        addLog "SilabasFinal:          " & ObjDTO.SilabasFinal
        addLog
        addLog "-------------------------------"
        addLog
        addLog "SilabaTonica:          " & ObjDTO.SilabasTonicas
        addLog "SilabaSecundaria:      " & ObjDTO.SilabasSecundarias
        addLog
        addLog "TextoFinal (idFonemas): " & ObjDTO.IdsFonemas
        addLog
        addLog
        addLog " (NOTA: Esta salida es ANSI, los caracterees '?'" & vbCrLf & "corresponden a fonemas que no se pueden representar)."
        addLog
        addLog "TextoFinal (Fonemas):  " & ObjDTO.FonemasFinal
        addLog
        addLog "-------------------------------"
        addLog "   Detalles internos"
        addLog "-------------------------------"
    
        addLog "Num sílabas auto:      " & CountItems(ObjDTO.SilabasAuto, " | ") + 1
        addLog "Num sílabas final:     " & CountItems(ObjDTO.SilabasFinal, " | ") + 1
        
        addLog
        addLog "Tiempo de proceso:     " & ObjDTO.Tiempo & " segundos"
        addLog "==============================="
        addLog   'vbCrLf
        
        PrintLog
    Else
        If DebugMotor Then
            PrintLog
        End If
    End If
        'Stop
    
    
'    Debug.Print strDebug
    
'    Open CurrentProject.Path & "\Debug.txt" For Output As #1
'    Print #1, strDebug
'    Close (1)
'
'    Shell "explorer " & CurrentProject.Path & "\Debug.txt"
    
End Sub

' ============================================================
' Contador auxiliar para separar elementos
' ============================================================
Private Function CountItems(ByVal s As String, ByVal sep As String) As Long
    If Len(Trim$(s)) = 0 Then
        CountItems = 0
    Else
        CountItems = UBound(Split(s, sep))
    End If
End Function


Public Sub addLog(Optional strlog As String = "")

    strDebug = strDebug & strlog & vbCrLf
    
End Sub


Public Sub PrintLog()

    Open CurrentProject.Path & "\Debug.txt" For Output As #1
    Print #1, strDebug
    Close (1)

    Shell "explorer " & CurrentProject.Path & "\Debug.txt"
    
End Sub

Public Sub InitLog()

    strDebug = ""
    
End Sub


Sub probar()

Dim t As String

t = LCase$("Almenarbrell")

Dim idx As Long
For idx = 1 To Len(t)
    Debug.Print idx, Mid$(t, idx, 1), Asc(Mid$(t, idx, 1))
Next idx

End Sub
