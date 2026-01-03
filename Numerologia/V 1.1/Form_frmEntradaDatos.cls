VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEntradaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private mIdiomaNom As clsIdioma
Private mIdiomaApe1 As clsIdioma
Private mIdiomaApe2 As clsIdioma

Public Property Get IdiomaNom() As clsIdioma
    Set IdiomaNom = mIdiomaNom
End Property

Public Property Set IdiomaNom(ByVal obj As clsIdioma)
    Set mIdiomaNom = obj
End Property

Public Property Get IdiomaApe1() As clsIdioma
    Set IdiomaApe1 = mIdiomaApe1
End Property

Public Property Set IdiomaApe1(ByVal obj As clsIdioma)
    Set mIdiomaApe1 = obj
End Property

Public Property Get IdiomaApe2() As clsIdioma
    Set IdiomaApe2 = mIdiomaApe2
End Property

Public Property Set IdiomaApe2(ByVal obj As clsIdioma)
    Set mIdiomaApe2 = obj
End Property



Private Sub cmdConvertir_Click()
    On Error GoTo ErrHandler
    
    Dim conv As New clsConversorFonetico
    Dim idiomaSel As idioma
    Dim nombreCompleto As String
    
    nombreCompleto = Me.txtNombreBase & " " & Me.txtApeBase1 & " " & Me.txtApeBase2

    ' Quitar espacios dobles interiores
    While InStr(nombreCompleto, "  ") > 0
        nombreCompleto = Replace(nombreCompleto, "  ", " ")
    Wend
    
    ' Quitar espacios al inicio y al final
    nombreCompleto = Trim(nombreCompleto)
    
    If nombreCompleto = "" Then
        MsgBox "Introduce al menos un nombre o apellido.", vbInformation, "Conversor Fonético"
        Exit Sub
    End If
    
    ' 2. Obtener idioma
    idiomaSel = Me.fraIdioma.Value
    
    ' 3. Configurar conversor
    conv.NombreOriginal = nombreCompleto
    conv.idioma = idiomaSel
    conv.UsarHmuda = Me.chkHmuda.Value
    conv.UsarUmuda = Me.chkUmuda.Value
    
    ' 4. Convertir
    Me.txtFonetico = conv.Convertir
    Me.txtNormal = nombreCompleto
    
    Exit Sub

ErrHandler:
    MsgBox "Error en la conversión fonética: " & Err.Description, vbExclamation, "Conversor Fonético"
End Sub

Private Sub cmdNom_Click()

    Dim nuevo As clsIdioma
    Set nuevo = SeleccionarIdioma("Nom", Me.IdiomaNom)

    If Not nuevo Is Nothing Then
        Set Me.IdiomaNom = nuevo
        Call CambiarIcono(Me.cmdNom, nuevo.Abreviado)
    End If

End Sub

Private Sub cmdIdiomaApe1_Click()

    Dim nuevo As clsIdioma
    Set nuevo = SeleccionarIdioma("Ape1", Me.IdiomaApe1)

    If Not nuevo Is Nothing Then
        Set Me.IdiomaApe1 = nuevo
        Call CambiarIcono(Me.cmdApe1, nuevo.Abreviado)
    End If

End Sub

Private Sub cmdIdiomaApe2_Click()

    Dim nuevo As clsIdioma
    Set nuevo = SeleccionarIdioma("Ape2", Me.IdiomaApe2)

    If Not nuevo Is Nothing Then
        Set Me.IdiomaApe2 = nuevo
        Call CambiarIcono(Me.cmdApe2, nuevo.Abreviado)
    End If

End Sub

Private Sub Form_Load()
    Me.fraSistema = 2
    
    
    ActualizarOpcionesSistema
    ActualizarOpcionesIdioma
End Sub

Private Sub fraSistema_Click()
    ActualizarOpcionesSistema
End Sub

Private Sub fraIdioma_Click()
    ActualizarOpcionesIdioma
End Sub

Private Sub ActualizarOpcionesSistema()
    With Me
        Dim activar As Boolean
        activar = (.fraSistema = 2)
        
        .chkHmuda.Enabled = activar
        .chkHmuda.Value = activar
        
        .chkUmuda.Enabled = activar
        .chkUmuda.Value = activar
    End With
End Sub
    
Private Sub ActualizarOpcionesIdioma()
'    With Me
'        Select Case .fraIdioma
'
'            Case 1, 4  ' Castellano, Galego
'                .chkHmuda.Enabled = True
'                .chkUmuda.Enabled = True
'
''            Case 4   ' Galego
''                .chkHmuda.Enabled = True
''                .chkUmuda.Enabled = True
'
'            Case 2   ' Catalán
'                .chkHmuda.Enabled = False
'                .chkHmuda.Value = True      ' H siempre muda
'                .chkUmuda.Enabled = False
'                .chkUmuda.Value = True      ' U muda en QU/GU
'
'            Case 3   ' Euskera
'                .chkHmuda.Enabled = False
'                .chkHmuda.Value = False     ' H siempre suena
'                .chkUmuda.Enabled = False
'                .chkUmuda.Value = False     ' U nunca muda
'
'        End Select
'    End With
End Sub

Public Sub CambiarIcono(btn As CommandButton, Abrev As String)
    Dim ruta As String
    ruta = CurrentProject.Path & "\img\" & Abrev & ".png"

    If Dir(ruta) = "" Then
        ruta = CurrentProject.Path & "\img\other.png"
    End If

    btn.Picture = ruta
End Sub


Private Function SeleccionarIdioma(campo As String, IdiomaActual As clsIdioma) As clsIdioma

    ' 1. Abrir oculto
    DoCmd.OpenForm "frmSelIdioma", , , , , acHidden

    ' 2. Pasar propiedades
    With Forms!frmSelIdioma
        .CampoDestino = campo
        Set .IdiomaActual = IdiomaActual
        .Modal = True
        .Visible = True
    End With

    ' 3. Al volver, recoger resultado
    If Forms!frmSelIdioma.Tag <> "" Then
        Set SeleccionarIdioma = Forms!frmSelIdioma.IdiomaSeleccionado
    Else
        Set SeleccionarIdioma = IdiomaActual   ' No hubo cambios
    End If

End Function


'Private Sub cmdConvertir_Click()
'    On Error GoTo ErrHandler
'
'    Dim conv As New clsConversorFonetico
'    Dim idiomaSel As Idioma
'
'    ' 1. Validar nombre
'    If Nz(Me.txtNombre.Value, "") = "" Then
'        MsgBox "Introduce un nombre para convertirlo a fonético.", vbInformation, "Conversor Fonético"
'        Exit Sub
'    End If
'
'    ' 2. Obtener idioma desde el cuadro de opciones fraIdioma
'    '    Asumiendo que las opciones tienen los valores de la enumeración Idioma:
'    '    1 = Español, 2 = Catalán, 3 = Euskera, 4 = Galego
'    idiomaSel = Me.fraIdioma.Value
'
'    ' 3. Configurar conversor
'    conv.NombreOriginal = Me.txtNombre.Value
'    conv.Idioma = idiomaSel
'    conv.UsarHmuda = Me.chkHmuda.Value
'    conv.UsarUmuda = Me.chkUmuda.Value
'
'    ' 4. Ejecutar conversión
'    Dim resultado As String
'    resultado = conv.Convertir
'
'    ' 5. Mostrar resultado
'    Me.txtFonetico.Value = resultado
''    me.txtNormal =
'
'    Exit Sub
'
'ErrHandler:
'    MsgBox "Error en la conversión fonética: " & Err.Description, vbExclamation, "Conversor Fonético"
'End Sub

