Attribute VB_Name = "modPropiedades"

Option Compare Database
Option Explicit


Private dbs As DAO.Database
Private prp As DAO.Property
Private prps As DAO.Properties

Public Sub SetProperty(strName As String, lngType As Long, _
   varValue As Variant)
'Created by Helen Feddema 31-Mar-2017
'Modified by Helen Feddema 31-Mar-2017
'Called from various procedures


On Error GoTo ErrorHandler
   'Attempt to set the specified property
   Set dbs = CurrentDb
   Set prps = dbs.Properties
   prps(strName) = varValue


ErrorHandlerExit:
   Exit Sub
ErrorHandler:
    If Err.Number = 3270 Then
      'The property was not found; create it
      Set prp = dbs.CreateProperty(Name:=strName, _
         Type:=lngType, Value:=varValue)
      CurrentDb.Properties.Append prp
      Resume Next
   Else
    MsgBox "Error No: " & Err.Number _
      & " in SetProperty procedure; " _
      & "Description: " & Err.Description
      Resume ErrorHandlerExit
   End If
End Sub


Public Function GetProperty(strName As String, strDefault As String) _
   As Variant
'Created by Helen Feddema 31-Mar-2017
'Modified by Helen Feddema 31-Mar-2017
'Called from various procedures
On Error GoTo ErrorHandler
   
   'Attempt to get the value of the specified property
   Set dbs = CurrentDb
   GetProperty = dbs.Properties(strName).Value
ErrorHandlerExit:
   Exit Function
ErrorHandler:
   If Err.Number = 3270 Then
      'The property was not found; use default value
      GetProperty = strDefault
      Resume Next
   Else
      MsgBox "Error No: " & Err.Number _
         & " in GetProperty procedure; " _
         & "Description: " & Err.Description
      Resume ErrorHandlerExit
   End If
End Function


Public Function ListCustomProps()
'Created by Helen Feddema 31-Mar-2017
'Modified by Helen Feddema 31-Mar-2017
'Lists custom database properties
On Error Resume Next
   
   Set dbs = CurrentDb
   Debug.Print "Custom database properties:"
   
   For Each prp In _
      dbs.Containers("Databases").Documents("UserDefined").Properties
      Debug.Print vbTab & prp.Name & ": " & prp.Value
   Next prp
End Function


Public Function ListAllProps()
'Created by Helen Feddema 31-Mar-2017
'Modified by Helen Feddema 31-Mar-2017
'Lists all database properties
On Error Resume Next
   
   Set dbs = CurrentDb
   Debug.Print "All database properties:"
   
   For Each prp In dbs.Properties
      Debug.Print vbTab & prp.Name & ": " & prp.Value
   Next prp
End Function

'========================================================================================================

Option Compare Database
Option Explicit


Private blnValue As Boolean
Private curValue As Currency
Private dblValue As Double
Private dteValue As Date
Private intValue As Integer
Private lngDataType As Long
Private lngValue As Currency
Private sglValue As Single
Private strFolderPath As String
Private strPrompt As String
Private strPropertyName As String
Private strPropertyValue As String
Private strText As String
Private strTitle As String

Private Sub chkBooleanValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   blnValue = Nz(Me![chkBooleanValue].Value, False)
   strPropertyName = "BooleanValue"
   lngDataType = dbBoolean
   Call SetProperty(strPropertyName, lngDataType, blnValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub cmdSelectFolderPath_Click()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   Dim fd As Office.FileDialog
   Dim txt As Access.TextBox
   
   'Create a FileDialog object as a Folder Picker dialog box.
   Set fd = Application.FileDialog(msoFileDialogFolderPicker)
   Set txt = Me![txtFolderPath]
   strPropertyName = "FolderPath"
   
   With fd
      .Title = "Browse for folder"
      .ButtonName = "Select"
      .InitialView = msoFileDialogViewDetails
      If .Show = -1 Then
         strPropertyValue = CStr(fd.SelectedItems.item(1))
         Debug.Print "Property value: " & strPropertyValue
         lngDataType = dbText
         Call SetProperty(strPropertyName, lngDataType, _
            strPropertyValue)
         txt.Value = strPropertyValue
      Else
         Debug.Print "User pressed Cancel"
      End If
   End With
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub Form_Load()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error Resume Next
   DoCmd.RunCommand acCmdSizeToFitForm
On Error GoTo ErrorHandler
   
   'Load control values from custom properties
   strFolderPath = Nz(GetProperty("FolderPath", ""))
   Me![txtFolderPath].Value = strFolderPath
   strText = Nz(GetProperty("TextValue", ""))
   Me![txtTextValue].Value = strText
   curValue = Nz(GetProperty("CurrencyValue", 0))
   Me![txtCurrencyValue].Value = curValue
   dteValue = Nz(GetProperty("DateValue", Date))
   Me![txtDateValue].Value = dteValue
   intValue = Nz(GetProperty("IntegerValue", 0))
   Me![txtIntegerValue].Value = intValue
   lngValue = Nz(GetProperty("LongValue", 0))
   Me![txtLongValue].Value = lngValue
   blnValue = Nz(GetProperty("BooleanValue", False))
   Me![chkBooleanValue].Value = blnValue
   dblValue = Nz(GetProperty("DoubleValue", 0))
   Me![txtDoubleValue].Value = dblValue
   sglValue = Nz(GetProperty("SingleValue", 0))
   Me![txtSingleValue].Value = sglValue
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.Name & " Form_Load procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtCurrencyValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   curValue = Nz(Me![txtCurrencyValue].Value, 0)
   strPropertyName = "CurrencyValue"
   lngDataType = dbCurrency
   Call SetProperty(strPropertyName, lngDataType, curValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtDateValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   If IsDate(Me![txtDateValue].Value) = False Then
      strTitle = "Invalid date"
      strPrompt = "Please enter a valid start date"
      Me![txtDateValue].SetFocus
      MsgBox Prompt:=strPrompt, _
         Buttons:=vbExclamation + vbOKOnly, _
         Title:=strTitle
      GoTo ErrorHandlerExit
   Else
      dteValue = Nz(Me![txtDateValue].Value, Date)
   End If
   
   strPropertyName = "DateValue"
   lngDataType = dbDate
   Call SetProperty(strPropertyName, lngDataType, dteValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtDoubleValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   dblValue = Nz(Me![txtDoubleValue].Value, 0)
   strPropertyName = "DoubleValue"
   lngDataType = dbDouble
   Call SetProperty(strPropertyName, lngDataType, dblValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtIntegerValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   intValue = Nz(Me![txtIntegerValue].Value, 0)
   strPropertyName = "IntegerValue"
   lngDataType = dbInteger
   Call SetProperty(strPropertyName, lngDataType, intValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtLongValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   lngValue = Nz(Me![txtLongValue].Value, 0)
   strPropertyName = "LongValue"
   lngDataType = dbLong
   Call SetProperty(strPropertyName, lngDataType, lngValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtSingleValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   sglValue = Nz(Me![txtSingleValue].Value, 0)
   strPropertyName = "SingleValue"
   lngDataType = dbSingle
   Call SetProperty(strPropertyName, lngDataType, sglValue)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub


Private Sub txtTextValue_AfterUpdate()
'Created by Helen Feddema 31-Mar-2017
'Last modified by Helen Feddema 31-Mar-2017


On Error GoTo ErrorHandler
   strText = Nz(Me![txtTextValue].Value, " ")
   strPropertyName = "TextValue"
   lngDataType = dbText
   Call SetProperty(strPropertyName, lngDataType, strText)
   
ErrorHandlerExit:
   Exit Sub
ErrorHandler:
   MsgBox "Error No: " & Err.Number _
      & " in " & Me.ActiveControl.Name & " procedure; " _
      & "Description: " & Err.Description
   Resume ErrorHandlerExit
End Sub
