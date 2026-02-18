Attribute VB_Name = "modMotor_Idioma_AUX_PT"
Option Compare Database
Option Explicit

Public Function EsVocal_PT(c As String) As Boolean
    EsVocal_PT = InStr("AEIOUÁÉÍÓÚÂÊÔaeiouáéíóúâêô", c) > 0
End Function

Public Function EsConsonant_PT(c As String) As Boolean
    EsConsonant_PT = Not EsVocal_PT(c) And c <> " "
End Function

Public Function EsGrupInseparable_PT(par As String) As Boolean
    Select Case UCase$(par)
        Case "BR", "BL", "CR", "CL", "DR", "TR", "PR", "PL", "GR", "GL", "FR", "FL"
            EsGrupInseparable_PT = True
        Case Else
            EsGrupInseparable_PT = False
    End Select
End Function

