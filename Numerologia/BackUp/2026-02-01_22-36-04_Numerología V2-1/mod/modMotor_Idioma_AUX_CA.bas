Attribute VB_Name = "modMotor_Idioma_AUX_CA"
Option Compare Database
Option Explicit

Public Function EsVocal_CA(c As String) As Boolean
    EsVocal_CA = InStr("AEIOUÀÈÉÍÏÒÓÚÜ", c) > 0
End Function

Public Function EsConsonant_CA(c As String) As Boolean
    EsConsonant_CA = Not EsVocal_CA(c) And c <> " " And c <> "·"
End Function

Public Function EsGrupInseparable_CA(par As String) As Boolean
    Select Case par
        Case "BR", "BL", "CR", "CL", "DR", "TR", "PR", "PL", "GR", "GL", "FR", "FL"
            EsGrupInseparable_CA = True
        Case Else
            EsGrupInseparable_CA = False
    End Select
End Function

'---------------------------------------------------------------------------------------------

Public Function EsDiptong_CA(v1 As String, v2 As String) As Boolean

    Dim d As String
    d = UCase$(v1 & v2)

    Select Case d
        Case "AI", "EI", "OI", "UI", _
             "AU", "EU", "OU", _
             "IA", "IE", "IO", "IU", _
             "UA", "UE", "UO"
            EsDiptong_CA = True
        Case Else
            EsDiptong_CA = False
    End Select

End Function

Public Function EsHiat_CA(v1 As String, v2 As String) As Boolean

    ' Vocal débil tónica ? siempre hiato
    If v1 = "Í" Or v1 = "Ú" Or v2 = "Í" Or v2 = "Ú" Then
        EsHiat_CA = True
        Exit Function
    End If

    ' Vocal fuerte + vocal fuerte ? hiato
    If InStr("AÀÁEÈÉOÒÓaàáeèéoòó", v1) > 0 And _
       InStr("AÀÁEÈÉOÒÓaàáeèéoòó", v2) > 0 Then
        EsHiat_CA = True
        Exit Function
    End If

    EsHiat_CA = False

End Function


