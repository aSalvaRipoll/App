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

''Public IdiomaActual As clsIdioma ' <-- ESTA es la que recibirá el valor
'Private mInicializado As Boolean ' <-- Para evitar recargar en cada Activate
''Public CampoDestino As String
''Public IdiomaSeleccionado As clsIdioma  '<-- Esta devuelve el valor

Dim obj As clsIdioma
Dim i As Long


'' Qué campo estamos editando: "Nom", "Ape1", "Ape2"
'Private mCampoDestino As String
'
'' Idioma actual (objeto)
'Private mIdiomaActual As clsIdioma
'
'' Idioma seleccionado (objeto)
'Private mIdiomaSeleccionado As clsIdioma
'
'
'Public Property Let CampoDestino(ByVal valor As String)
'    mCampoDestino = valor
'End Property
'
'Public Property Get CampoDestino() As String
'    CampoDestino = mCampoDestino
'End Property
'
'Public Property Set IdiomaActual(ByVal obj As clsIdioma)
'    Set mIdiomaActual = obj
'End Property
'
'Public Property Get IdiomaActual() As clsIdioma
'    Set IdiomaActual = mIdiomaActual
'End Property
'
'Public Property Set IdiomaSeleccionado(ByVal obj As clsIdioma)
'    Set mIdiomaSeleccionado = obj
'End Property
'
'Public Property Get IdiomaSeleccionado() As clsIdioma
'    Set IdiomaSeleccionado = mIdiomaSeleccionado
'End Property

'Private Sub Form_Activate()
''
''    Static hecho As Boolean
'    Dim obj As clsIdioma
'    Dim i As Long
'
''    If hecho Then Exit Sub
''    hecho = True
''
''    ' Si no te han pasado idioma, pon uno por defecto
'''    If mIdiomaActual Is Nothing Then
'''        Set mIdiomaActual = colIdiomas("1")   ' Español, por ejemplo
'''    End If
''
''    If IdiomaActual Is Nothing Then
''        Set IdiomaActual = colIdiomas("0")   ' Español, por ejemplo
''    End If
''
''    ' Mostrar nombre actual
''    Me.lblIdiomaActual.Caption = IdiomaActual.NomIdioma
''
''    ' Cargar bandera actual
''    Me.imgActual.Picture = CurrentProject.Path & "\img\" & IdiomaActual.Abreviado & ".ico"
'
''    ' Cargar lista
''    Me.lstIdiomas.RowSource = ""
''    For Each obj In colIdiomas
''        Me.lstIdiomas.AddItem obj.IDIdioma & ";" & obj.Abreviado & ";" & obj.NomIdioma
''    Next obj
''
''    Me.lstIdiomas.Requery
''
''    ' Seleccionar el idioma actual
''    For i = 0 To Me.lstIdiomas.ListCount - 1
''        If Me.lstIdiomas.Column(0, i) = IdiomaActual.IDIdioma Then
''            Me.lstIdiomas.Selected(i) = True
''            Exit For
''        End If
''    Next i
'
'End Sub

Private Sub Form_Current()

   ' Cargar lista
    Me.lstIdiomas.RowSource = ""
    For Each obj In colIdiomas
        Me.lstIdiomas.AddItem obj.IDIdioma & ";" & obj.Abreviado & ";" & obj.NomIdioma
    Next obj

    Me.lstIdiomas.Requery
    
    ' Seleccionar el idioma actual
    For i = 0 To Me.lstIdiomas.ListCount - 1
        If Me.lstIdiomas.Column(0, i) = IdiomaActual.IDIdioma Then
            Me.lstIdiomas.Selected(i) = True
            Exit For
        End If
    Next i
    
    lstIdiomas_Click
    
End Sub

'Private Sub Form_Load()
'
'    If colIdiomas Is Nothing Then
'        CargarIdiomas
'    End If
'
'End Sub


Private Sub Form_Load()



    If colIdiomas Is Nothing Then CargarIdiomas
    If IdiomaActual Is Nothing Then Set IdiomaActual = colIdiomas("0")

    Me.lblIdiomaActual.Caption = IdiomaActual.NomIdioma
    Me.imgActual.Picture = CurrentProject.Path & "\img\" & IdiomaActual.Abreviado & ".ico"
    'Me.Tag = CampoDestino

    If IdiomaActual Is Nothing Then
        Me.lblIdiomaActual.Caption = "(sin idioma)"
'        Exit Sub
    End If

    ' Mostrar nombre actual
    Me.lblIdiomaActual.Caption = IdiomaActual.NomIdioma

    ' Cargar bandera actual
    'Me.imgActual.Picture = CurrentProject.Path & "\img\" & mIdiomaActual.Abreviado & ".png"
    Me.imgActual.Picture = CurrentProject.Path & "\img\" & IdiomaActual.Abreviado & ".ico"

    ' Cargar lista
    If colIdiomas Is Nothing Then
        Call CargarIdiomas
    End If

'    Me.lstIdiomas.RowSource = ""
'    For Each obj In colIdiomas
'        Me.lstIdiomas.AddItem obj.IDIdioma & ";" & obj.Abreviado & ";" & obj.NomIdioma
'    Next obj
'
'    Me.lstIdiomas.Requery
'
'    ' Seleccionar el idioma actual
'    For i = 0 To Me.lstIdiomas.ListCount - 1
'        If Me.lstIdiomas.Column(0, i) = IdiomaActual.IDIdioma Then
'            Me.lstIdiomas.Selected(i) = True
'            Exit For
'        End If
'    Next i

End Sub


'Private Sub lstIdiomas_AfterUpdate()
'
'    Dim ID As Byte
'    Dim obj As clsIdioma
'
'    ID = Me.lstIdiomas.Column(0)
'
'    Set obj = colIdiomas(CStr(ID))
'
'    'Me.imgNueva.Picture = CurrentProject.Path & "\img\" & obj.Abreviado & ".png"
'    Me.imgNueva.Picture = CurrentProject.Path & "\img\" & obj.Abreviado & ".ico"
'
'    Set IdiomaSeleccionado = obj
'
'End Sub

Private Sub cmdAceptar_Click()

    If IdiomaSeleccionado Is Nothing Then
        Set IdiomaSeleccionado = IdiomaActual
    End If

    Me.Tag = CampoDestino

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdCancelar_Click()
    Me.Tag = ""
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub lstIdiomas_Click()

    Dim ID As Byte
    Dim obj As clsIdioma

    ID = Me.lstIdiomas.Column(0)

    Set obj = colIdiomas(CStr(ID))

    'Me.imgNueva.Picture = CurrentProject.Path & "\img\" & obj.Abreviado & ".png"
    Me.imgNueva.Picture = CurrentProject.Path & "\img\" & obj.Abreviado & ".ico"

    Set IdiomaSeleccionado = obj

End Sub
