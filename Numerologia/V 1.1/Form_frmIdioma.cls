VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Qué campo estamos editando: "Nom", "Ape1", "Ape2"
Private mCampoDestino As String

' Idioma actual (objeto)
Private mIdiomaActual As clsIdioma

' Idioma seleccionado (objeto)
Private mIdiomaSeleccionado As clsIdioma


Public Property Let CampoDestino(ByVal valor As String)
    mCampoDestino = valor
End Property

Public Property Get CampoDestino() As String
    CampoDestino = mCampoDestino
End Property

Public Property Set IdiomaActual(ByVal obj As clsIdioma)
    Set mIdiomaActual = obj
End Property

Public Property Get IdiomaActual() As clsIdioma
    Set IdiomaActual = mIdiomaActual
End Property

Public Property Set IdiomaSeleccionado(ByVal obj As clsIdioma)
    Set mIdiomaSeleccionado = obj
End Property

Public Property Get IdiomaSeleccionado() As clsIdioma
    Set IdiomaSeleccionado = mIdiomaSeleccionado
End Property


Private Sub Form_Load()

    Dim obj As clsIdioma
    Dim i As Long

    If mIdiomaActual Is Nothing Then
        Me.lblIdiomaActual.Caption = "(sin idioma)"
        Exit Sub
    End If

    ' Mostrar nombre actual
    Me.lblIdiomaActual.Caption = mIdiomaActual.NomIdioma

    ' Cargar bandera actual
    Me.imgActual.Picture = CurrentProject.Path & "\img\" & mIdiomaActual.Abreviado & ".png"

    ' Cargar lista
    Me.lstIdiomas.RowSource = ""
    For Each obj In colIdiomas
        Me.lstIdiomas.AddItem obj.IDIdioma & ";" & obj.Abreviado & ";" & obj.NomIdioma
    Next obj

    ' Seleccionar el idioma actual
    For i = 0 To Me.lstIdiomas.ListCount - 1
        If Me.lstIdiomas.Column(0, i) = mIdiomaActual.IDIdioma Then
            Me.lstIdiomas.Selected(i) = True
            Exit For
        End If
    Next i

End Sub


Private Sub lstIdiomas_AfterUpdate()

    Dim ID As Byte
    Dim obj As clsIdioma

    ID = Me.lstIdiomas.Column(0)

    Set obj = colIdiomas(CStr(ID))

    Me.imgNueva.Picture = CurrentProject.Path & "\img\" & obj.Abreviado & ".png"

    Set mIdiomaSeleccionado = obj

End Sub

Private Sub cmdAceptar_Click()

    If mIdiomaSeleccionado Is Nothing Then
        Set mIdiomaSeleccionado = mIdiomaActual
    End If

    Me.Tag = mCampoDestino

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdCancelar_Click()
    Me.Tag = ""
    DoCmd.Close acForm, Me.Name
End Sub




