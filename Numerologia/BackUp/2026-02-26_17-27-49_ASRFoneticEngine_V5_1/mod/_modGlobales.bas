Attribute VB_Name = "_modGlobales"

Option Compare Database
Option Explicit

' _modGlobales
' Módulo que contiene todos los elementos globales de la aplicación

Public Type tAppVersion
    vMajor As Integer
    vMinor As Integer
    vVersion As Integer
End Type

Public AppVersion As tAppVersion


Sub InitApp()
    With AppVersion
        .vMajor = GetProperty("vMajor", 0)
        .vMinor = GetProperty("vMinor", 0)
        .vVersion = GetProperty("vVersion", 0)
    End With
End Sub


Public Function GetProperty(strName As String, strDefault As String) _
   As Variant
   
   Dim dbs As Object
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


