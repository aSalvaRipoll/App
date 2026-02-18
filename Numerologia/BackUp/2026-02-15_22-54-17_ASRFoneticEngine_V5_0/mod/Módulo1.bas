Attribute VB_Name = "Módulo1"
Option Compare Database
Option Explicit



Sub test()


#If VBA7 Then
    Debug.Print "VBA7"
#Else
    Debug.Print "VBA6"
#End If

End Sub

