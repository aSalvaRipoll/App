Attribute VB_Name = "modSilabeoAuxiliares"

' ============================================================
'  MÓDULO: modSilabeoAuxiliares
'  Funciones auxiliares comunes a TODOS los idiomas
' ============================================================

Option Compare Text
Option Explicit

' ------------------------------------------------------------
' 1. EsMiembro: comprueba si una cadena está en un array
' ------------------------------------------------------------
Public Function EsMiembro(ByVal s As String, ByVal arr As Variant) As Boolean
    Dim x As Variant
    For Each x In arr
        If s = x Then
            EsMiembro = True
            Exit Function
        End If
    Next x
    EsMiembro = False
End Function

' ------------------------------------------------------------
' 2. EsVocal: vocal universal (para todos los idiomas)
' ------------------------------------------------------------
Public Function EsVocal(ByVal c As String) As Boolean
    EsVocal = InStr("AEIOUÁÉÍÓÚÀÈÌÒÙÂÊÔÃÕÜ", c) > 0
End Function

' ------------------------------------------------------------
' 3. MF_UnirConsonantesEnAtaque
'    Une dos sílabas cuando el universal las separó
' ------------------------------------------------------------
Public Sub MF_UnirConsonantesEnAtaque( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        ' Si la división está justo antes de pos ? unir
        If fin = pos - 1 Then
            Dim ini2 As Long, fin2 As Long
            ini2 = silabas(i + 1)(0)
            fin2 = silabas(i + 1)(1)

            silabas.Remove i + 1
            silabas.Remove i
            silabas.Add Array(ini, fin2), , i
            Exit Sub
        End If
    Next i

End Sub

' ------------------------------------------------------------
' 4. MF_UnirVocalesEnDiptongo
'    Une dos sílabas cuando deben formar diptongo
' ------------------------------------------------------------
Public Sub MF_UnirVocalesEnDiptongo( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        ' Si la sílaba empieza en pos ? unir con la anterior
        If ini = pos Then
            Dim iniPrev As Long, finPrev As Long
            iniPrev = silabas(i - 1)(0)
            finPrev = silabas(i - 1)(1)

            silabas.Remove i
            silabas.Remove i - 1
            silabas.Add Array(iniPrev, fin), , i - 1
            Exit Sub
        End If
    Next i

End Sub

' ------------------------------------------------------------
' 5. MF_ForzarDivisionSilabica
'    Divide una sílaba en dos en un punto concreto
' ------------------------------------------------------------
Public Sub MF_ForzarDivisionSilabica( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        If pos > ini And pos <= fin Then
            ' Dividir esta sílaba en dos
            silabas.Remove i
            silabas.Add Array(ini, pos - 1), , i
            silabas.Add Array(pos, fin), , i + 1
            Exit Sub
        End If
    Next i

End Sub

' ------------------------------------------------------------
' 6. MF_EliminarVocalFinal
'    (para francés: elimina la E muda final)
' ------------------------------------------------------------
Public Sub MF_EliminarVocalFinal( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        If fin = pos Then
            silabas.Remove i
            silabas.Add Array(ini, pos - 1), , i
            Exit Sub
        End If
    Next i

End Sub


