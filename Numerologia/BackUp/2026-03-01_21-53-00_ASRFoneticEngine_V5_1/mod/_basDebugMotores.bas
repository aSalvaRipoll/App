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
Public Sub MF_DebugDTO(proc As String)

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
        addLog "     CONTENIDO DEL DTO"
        addLog "    (VARIABLES INTERNAS)"
        addLog "==============================="
    
        addLog
        addLog "-------------------------------"
        addLog " Proc.: " & proc
        addLog "-------------------------------"
        addLog
        
        addLog "Texto original (ObjDTO.TextoOriginal):        " & vbCrLf & vbTab & vbTab & "'" & ObjDTO.TextoOriginal & "'"
        addLog
        addLog "Texto Corregido (ObjDTO.TextoCorregido):       " & vbCrLf & vbTab & vbTab & "'" & ObjDTO.TextoCorregido & "'"
        addLog
        addLog "Texto normalizado (ObjDTO.TextoNormalizado):     " & vbCrLf & vbTab & vbTab & "'" & ObjDTO.TextoNormalizado & "'"
        addLog
        addLog "-------------------------------"
        addLog
        addLog "SilabasAuto (ObjDTO.SilabasAuto):           " & vbCrLf & vbTab & vbTab & ObjDTO.SilabasAuto
        addLog
        addLog "SilabasAcentuadas (ObjDTO.SilabasAcentuadas):     " & vbCrLf & vbTab & vbTab & ObjDTO.SilabasAcentuadas
        addLog
        addLog "SilabasFinal (ObjDTO.SilabasFinal):          " & vbCrLf & vbTab & vbTab & ObjDTO.SilabasFinal
        addLog
        addLog "-------------------------------"
        addLog
        addLog "Silaba(s) Tonica(s) (ObjDTO.SilabasTonicas):          " & vbCrLf & vbTab & vbTab & ObjDTO.SilabasTonicas
        addLog "SilabaSecundaria (ObjDTO.SilabasSecundarias):      " & vbCrLf & vbTab & vbTab & ObjDTO.SilabasSecundarias
        addLog
        addLog "Texto Final (idFonemas) (ObjDTO.IdsFonemas): " & vbCrLf & vbTab & vbTab & ObjDTO.IdsFonemas
        addLog
        addLog
        addLog " (NOTA: Esta salida es ANSI, los caracterees '?'" & vbCrLf & "corresponden a fonemas que no se pueden representar)."
        addLog
        addLog "Texto Final (Fonemas) (ObjDTO.FonemasFinal):  " & vbCrLf & vbTab & vbTab & ObjDTO.FonemasFinal
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
    
'    If DebugDTO Then
'        If Not DebugMotor Then
'            strDebug = ""
'        End If
'
'        If ObjDTO Is Nothing Then
'            addLog "DTO no inicializado."
'            Exit Sub
'        End If
'
'        addLog
'        addLog StrConv("===============================", vbUnicode)
'        addLog StrConv("     ESTADO DEL DTO", vbUnicode)
'        addLog StrConv("===============================", vbUnicode)
'
'        addLog
'        addLog StrConv("-------------------------------", vbUnicode)
'        addLog StrConv(" Proc.: " & Proc, vbUnicode)
'        addLog StrConv("-------------------------------", vbUnicode)
'        addLog
'
'        addLog StrConv("Texto original:        " & ObjDTO.TextoOriginal, vbUnicode)
'        addLog
'        addLog StrConv("Texto Corregido:       " & ObjDTO.TextoCorregido, vbUnicode)
'        addLog
'        addLog StrConv("Texto normalizado:     " & ObjDTO.TextoNormalizado, vbUnicode)
'        addLog
'        addLog StrConv("-------------------------------", vbUnicode)
'        addLog
'        addLog StrConv("SilabasAuto:           " & ObjDTO.SilabasAuto, vbUnicode)
'        addLog
'        addLog StrConv("SilabasAcentuadas:     " & ObjDTO.SilabasAcentuadas, vbUnicode)
'        addLog
'        addLog StrConv("SilabasFinal:          " & ObjDTO.SilabasFinal, vbUnicode)
'        addLog
'        addLog StrConv("-------------------------------", vbUnicode)
'        addLog
'        addLog StrConv("SilabaTonica:          " & ObjDTO.SilabasTonicas, vbUnicode)
'        addLog StrConv("SilabaSecundaria:      " & ObjDTO.SilabasSecundarias, vbUnicode)
'        addLog
'        addLog StrConv("TextoFinal (idFonemas): " & ObjDTO.IdsFonemas, vbUnicode)
'        addLog
'        addLog
'        addLog StrConv(" (NOTA: Esta salida es ANSI, los caracterees '?'" & vbCrLf & "corresponden a fonemas que no se pueden representar).", vbUnicode)
'        addLog
'        addLog StrConv("TextoFinal (Fonemas):  " & ObjDTO.FonemasFinal, vbUnicode)
'        addLog
'        addLog StrConv("-------------------------------", vbUnicode)
'        addLog StrConv("   Detalles internos", vbUnicode)
'        addLog StrConv("-------------------------------", vbUnicode)
'
'        addLog StrConv("Num sílabas auto:      " & CountItems(ObjDTO.SilabasAuto, " | ") + 1, vbUnicode)
'        addLog StrConv("Num sílabas final:     " & CountItems(ObjDTO.SilabasFinal, " | ") + 1, vbUnicode)
'
'        addLog
'        addLog StrConv("Tiempo de proceso:     " & ObjDTO.Tiempo & " segundos", vbUnicode)
'        addLog StrConv("===============================", vbUnicode)
'        addLog   'vbCrLf
'
'        PrintLog
'    Else
'        If DebugMotor Then
'            PrintLog
'        End If
'    End If
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

'    Dim fs As Scripting.FileSystemObject
'    Dim a As Object
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set a = fs.CreateTextFile(CurrentProject.Path & "\Debug.txt", True)
'
'    a.Write strDebug '("This is a test.")
'    a.Close

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
