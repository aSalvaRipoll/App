VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSelectorPersonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdAceptar_Click()

   Call lstPersonas_DblClick(False)

End Sub

Private Sub Comando8_Click()
    
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub Form_Load()

'------------------------------------------------------
' SQL del Listbox para agrgar las cuentas de fonéticas.
'------------------------------------------------------
'
'SELECT
'    P.ID_Persona,
'    P.Nombre & ' ' & P.Ape1 & ' ' & P.Ape2 AS NombreCompleto,
'    P.FechaNacimiento,
'    G.Genero,
'
'    Nz((SELECT Count(*)
'        FROM Fonetica
'        WHERE IDPersona = P.ID_Persona AND Activo = True), 0) AS FonActivas,
'
'    Nz((SELECT Count(*)
'        FROM Fonetica
'        WHERE IDPersona = P.ID_Persona AND Activo = False), 0) AS FonBorradas
'
'FROM
'    tbuPersonas As p
'    INNER JOIN tbmGeneros AS G ON P.ID_Genero = G.ID
'
'ORDER BY
'    P.Nombre, P.Ape1, P.Ape2;
'
'------------------------------------------------------

'Versión 2 con Iconos.

'SELECT
'    I.Icono,
'    P.ID_Persona,
'    P.Nombre & ' ' & P.Ape1 & ' ' & P.Ape2 AS NombreCompleto,
'    P.FechaNacimiento,
'    G.Genero,
'    Nz(FonActivas.Cnt,0) AS FonActivas,
'    Nz(FonBorradas.Cnt, 0) As FonBorradas
'
'FROM
'    tbuPersonas As p
'    INNER JOIN tbmGeneros AS G ON P.ID_Genero = G.ID
'
'    LEFT JOIN (
'        SELECT IDPersona, Count(*) AS Cnt
'        FROM Fonetica
'        WHERE Activo = True
'        GROUP BY IDPersona
'    ) AS FonActivas ON FonActivas.IDPersona = P.ID_Persona
'
'    LEFT JOIN (
'        SELECT IDPersona, Count(*) AS Cnt
'        FROM Fonetica
'        WHERE Activo = False
'        GROUP BY IDPersona
'    ) AS FonBorradas ON FonBorradas.IDPersona = P.ID_Persona
'
'    LEFT JOIN tbmIconos AS I
'        ON I.ID = IIf(Nz(FonActivas.Cnt,0) > 0, 1, 2)
'
'ORDER BY
'    P.Nombre, P.Ape1, P.Ape2;

End Sub

Private Sub lstPersonas_DblClick(Cancel As Integer)

    Dim ID As Long
    ID = Me.lstPersonas.Column(0)   ' primera columna = IDPersona

    ' 1. Cargar la clase directamente
    Call CargarClaseDesdeTabla(ID)

    ' 2. Pedir al formulario principal que refresque su interfaz
    Forms!frmEntradaDatos.RefrescarDesdeClase

    ' 3. Cerrar selector
    DoCmd.Close acForm, Me.Name

End Sub


Private Sub txtBuscar_Change()

    Dim filtro As String
    filtro = Me.txtBuscar.Text

    Me.lstPersonas.RowSource = _
        "SELECT " & _
        "tbuPersonas.ID_Persona, " & _
        "tbuPersonas.Nombre & ' ' & tbuPersonas.Ape1 & ' ' & tbuPersonas.Ape2 AS NombreCompleto, " & _
        "tbuPersonas.FechaNacimiento, " & _
        "tbmGeneros.Genero " & _
        "FROM tbuPersonas " & _
        "INNER JOIN tbmGeneros ON tbuPersonas.ID_Genero = tbmGeneros.ID " & _
        "WHERE tbuPersonas.Nombre LIKE '*" & filtro & "*' " & _
        "OR tbuPersonas.Ape1 LIKE '*" & filtro & "*' " & _
        "OR tbuPersonas.Ape2 LIKE '*" & filtro & "*' " & _
        "ORDER BY tbuPersonas.Nombre, tbuPersonas.Ape1, tbuPersonas.Ape2;"

End Sub


