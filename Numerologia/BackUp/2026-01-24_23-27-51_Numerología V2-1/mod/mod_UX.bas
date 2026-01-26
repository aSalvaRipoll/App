Attribute VB_Name = "mod_UX"

Option Compare Database
Option Explicit

' =====================================================
' Estilos generales
' =====================================================

Public Sub EstiloFormulario(frm As Form)

    Dim ctl As Control

    For Each ctl In frm.Controls

        Select Case ctl.ControlType

            Case acLabel
                Call EstiloLabel(ctl)
                
            Case acTextBox
                Call EstiloTextBox(ctl)

            Case acComboBox
                Call EstiloCombo(ctl)

            Case acCommandButton
                Call EstiloBoton(ctl)

'            Case acCheckBox
'                Call EstiloCheck(ctl)

'            Case acOptionButton
'                Call EstiloOption(ctl)

            Case acOptionGroup
                Call EstiloOptionGroup(ctl)

        End Select

    Next ctl

End Sub

Public Sub EstiloLabel(lbl As Label)
    On Error Resume Next
    lbl.FontName = "Segoe UI"
    lbl.FontSize = 10
    lbl.ForeColor = RGB(40, 40, 40)
End Sub

Public Sub EstiloTextBox(txt As TextBox)
    On Error Resume Next
    txt.FontName = "Segoe UI"
    txt.FontSize = 10
    txt.BackColor = RGB(255, 255, 255)
    txt.ForeColor = RGB(0, 0, 0)
End Sub

Public Sub EstiloCombo(cbo As ComboBox)
    On Error Resume Next
    cbo.FontName = "Segoe UI"
    cbo.FontSize = 10
End Sub

Public Sub EstiloBoton(btn As CommandButton)
    On Error Resume Next
    btn.FontName = "Segoe UI"
    btn.FontSize = 10
    btn.ForeColor = RGB(0, 70, 140)
End Sub

'Public Sub EstiloCheck(chk As CheckBox)
'    On Error Resume Next
'    chk.FontName = "Segoe UI"
'    chk.FontSize = 10
'End Sub

'Public Sub EstiloOption(opt As OptionButton)
'    On Error Resume Next
'    opt.FontName = "Segoe UI"
'    opt.FontSize = 10
'End Sub

Public Sub EstiloOptionGroup(grp As OptionGroup)
    On Error Resume Next
    grp.BackColor = RGB(240, 240, 240)
End Sub

' =====================================================

Public Sub EstiloTitulosAuto(frm As Form)

    Dim lblTitulo As Label
    Dim lblSombra As Label
    Dim ctl As Control

    ' Buscar lblTitulo1 y lblTitulo2
    For Each ctl In frm.Controls
        If TypeOf ctl Is Label Then
            If ctl.Name = "lblTitulo1" Then Set lblTitulo = ctl
            If ctl.Name = "lblTitulo2" Then Set lblSombra = ctl
        End If
    Next ctl

    ' Si ambos existen, aplicar estilo
    If Not lblTitulo Is Nothing And Not lblSombra Is Nothing Then
        Call EstiloTituloMain(lblTitulo, lblSombra)
        Call AlineaTituloCentro(lblTitulo, lblSombra)
    End If

End Sub

' =====================================================
' Resaltar controles
' =====================================================

Public Sub ResaltaControl(ctrl As Control, _
                          Optional ByVal ColorError As Long = vbYellow, _
                          Optional ByVal ColorNormal As Long = vbWhite, _
                          Optional ByVal Milisegundos As Long = 600)

    On Error Resume Next

    ctrl.BackColor = ColorError
    DoEvents

    Dim t As Single
    t = Timer

    ' Espera no bloqueante
    Do While Timer < t + (Milisegundos / 1000)
        DoEvents
    Loop

    ctrl.BackColor = ColorNormal
End Sub

' =====================================================
' Resalte con parpadeo
' =====================================================

Public Sub ParpadeaControl(ctrl As Control, _
                           Optional Veces As Integer = 2, _
                           Optional ColorError As Long = vbYellow, _
                           Optional ColorNormal As Long = vbWhite, _
                           Optional Intervalo As Long = 150)

    Dim i As Integer
    For i = 1 To Veces
        ctrl.BackColor = ColorError
        DoEvents
        Espera Intervalo

        ctrl.BackColor = ColorNormal
        DoEvents
        Espera Intervalo
    Next i
End Sub

Private Sub Espera(ms As Long)
    Dim t As Single
    t = Timer
    Do While Timer < t + (ms / 1000)
        DoEvents
    Loop
End Sub

' =====================================================
' Aplicar Estilo específico al título
' =====================================================

Public Sub EstiloTituloMain(lblTexto As Label, lblSombra As Label, _
                        Optional dx As Integer = 45, _
                        Optional dy As Integer = 45)

    On Error Resume Next

    ' Alinear Etiqueta Título
    AlineaTituloCentro lblTexto, lblSombra, dx

    ' Establecer estilo base
    lblTexto.FontName = "Segoe UI"
    lblTexto.FontBold = True
    lblTexto.ForeColor = RGB(0, 70, 140)
    lblTexto.TextAlign = 1

    ' Sincronizar estilo y texto
    lblSombra.Caption = lblTexto.Caption
    lblSombra.FontName = lblTexto.FontName
    lblSombra.FontSize = lblTexto.FontSize
    lblSombra.FontBold = lblTexto.FontBold
    lblSombra.TextAlign = lblTexto.TextAlign
    
    ' Color de la sombra
    lblSombra.ForeColor = RGB(60, 60, 60)
    'lblSombra.ForeColor = vbBlack

    
    ' Posición y tamaño
    lblSombra.Left = lblTexto.Left + dx
    lblSombra.Top = lblTexto.Top + dy

    lblSombra.width = lblTexto.width
    lblSombra.height = lblTexto.height
    
    ' Eliminar foco de la sombra
'    lblSombra.Enabled = False
'    lblSombra.Locked = True

    ' Orden de profundidad
'    lblSombra.ZOrder 1
'    lblTexto.ZOrder 0

End Sub

Public Sub AlineaTituloCentro(lblTexto As Label, lblSombra As Label, _
                              Optional dx As Integer = 45)

    Dim frm As Form
    Set frm = lblTexto.Parent

    lblTexto.Left = (frm.InsideWidth - lblTexto.width) \ 2
    lblSombra.Left = lblTexto.Left + dx

End Sub

' =====================================================

Public Sub AjustaTamanoTitulo(lbl As Label, AnchoMax As Long, _
                              Optional TamanoMax As Integer = 28, _
                              Optional TamanoMin As Integer = 12)

    Dim t As Integer
    t = TamanoMax

    lbl.FontSize = t

    Do While lbl.width > AnchoMax And t > TamanoMin
        t = t - 1
        lbl.FontSize = t
    Loop

End Sub

'Sub AlinearTexto()
'
''0 Print Izquierda
''1 Print Centro
''2 Print Derecha
''3 Print Justificado
''4 Print Distribuido
'
'TextAlign = 1
'
'End Sub

'Public Sub AlineaTituloCentro_old(frm As Form, lblTexto As Label, lblSombra As Label, _
'                              Optional dx As Integer = 45)
'
'    lblTexto.Left = (frm.InsideWidth - lblTexto.width) \ 2
'    lblSombra.Left = lblTexto.Left + dx
'
'End Sub






'Public Sub AplicaSombra(lblTexto As Label, lblSombra As Label, _
'                        Optional dx As Integer = 45, _
'                        Optional dy As Integer = 45)
'
'    On Error Resume Next
'
'    ' Sincronizar estilo
'    lblSombra.Caption = lblTexto.Caption
'    lblSombra.FontName = lblTexto.FontName
'    lblSombra.FontSize = lblTexto.FontSize
'    lblSombra.FontBold = lblTexto.FontBold
'
'    ' Color de la sombra
'    lblSombra.ForeColor = RGB(60, 60, 60)
'    'lblSombra.ForeColor = vbBlack
'
'    ' Posición y tamaño
'    lblSombra.Left = lblTexto.Left + dx
'    lblSombra.Top = lblTexto.Top + dy
'
'    lblSombra.width = lblTexto.width
'    lblSombra.height = lblTexto.height
'
'    ' Orden de profundidad
'    lblSombra.ZOrder 1
'    lblTexto.ZOrder 0
'
'End Sub
'
'Public Sub AplicaSombraBase(lblTexto As Label, lblSombra As Label, _
'                        Optional DesfaseX As Integer = 1, _
'                        Optional DesfaseY As Integer = 1)
'
'    On Error Resume Next
'
'    lblSombra.Caption = lblTexto.Caption
'    lblSombra.FontName = lblTexto.FontName
'    lblSombra.FontSize = lblTexto.FontSize
'    lblSombra.FontBold = lblTexto.FontBold
'    lblSombra.ForeColor = vbBlack
'
'    lblSombra.Left = lblTexto.Left + DesfaseX
'    lblSombra.Top = lblTexto.Top + DesfaseY
'
'End Sub
'
'
'Public Sub EstiloTitulo(lblTexto As Label, lblSombra As Label)
'
'    lblTexto.FontName = "Segoe UI"
'    lblTexto.FontBold = True
'    lblTexto.ForeColor = RGB(0, 70, 140)
'
'    lblSombra.Caption = lblTexto.Caption
'    lblSombra.FontName = lblTexto.FontName
'    lblSombra.FontSize = lblTexto.FontSize '= 20
'    lblSombra.FontBold = lblTexto.FontBold
'
'    Call AplicaSombra(lblTexto, lblSombra, 1, 1)
'End Sub



