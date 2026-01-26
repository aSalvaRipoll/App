Attribute VB_Name = "__GeneradorClases"

Option Compare Database
Option Explicit

Public Sub GenerarClaseDesdeTabla(NombreTabla As String, Optional NombreClase As String = "", Optional DTO As Boolean = False)

    Dim db As dao.Database
    Dim tdf As dao.TableDef
    Dim fld As dao.Field
'    Dim s As String
    Dim TipoVBA As String

    Set db = CurrentDb
    Set tdf = db.TableDefs(NombreTabla)

    If NombreClase = "" Then NombreClase = NombreTabla

    Debug.Print
    Debug.Print "' ===================================================================="
    Debug.Print "' Clase generada automáticamente desde la tabla: " & NombreTabla
    Debug.Print "' ===================================================================="
    Debug.Print
    Debug.Print "Option Compare Database"
    Debug.Print "Option Explicit"
    Debug.Print
    Debug.Print "' Public Class " & NombreClase
    Debug.Print


    If DTO = True Then
        ' ---------------------------------------------------------
        ' 1. Propiedades públicas
        ' ---------------------------------------------------------
    
        For Each fld In tdf.Fields
            TipoVBA = TipoAccessAVBA(fld.Type)
    
            Debug.Print "    Public " & fld.Name & " As " & TipoVBA & "' Campo: " & fld.Name
            'Debug.Print
            
        Next fld
        Debug.Print
        'Debug.Print "' ===================================================================="
        Debug.Print "' End Class"
        Debug.Print "' ===================================================================="

    Else
        ' ---------------------------------------------------------
        ' 1. Declaración de variables privadas
        ' ---------------------------------------------------------
        For Each fld In tdf.Fields
            TipoVBA = TipoAccessAVBA(fld.Type)
            Debug.Print "    Private m" & fld.Name & " As " & TipoVBA & "' Campo: " & fld.Name
        Next fld
    
        Debug.Print
    
        ' ---------------------------------------------------------
        ' 2. Propiedades públicas
        ' ---------------------------------------------------------
        For Each fld In tdf.Fields
            TipoVBA = TipoAccessAVBA(fld.Type)
    
            Debug.Print "    Public Property Get " & fld.Name & "() As " & TipoVBA
            Debug.Print "        " & fld.Name & " = m" & fld.Name
            Debug.Print "    End Property"
            Debug.Print
    
            Debug.Print "    Public Property Let " & fld.Name & "(ByVal Valor As " & TipoVBA & ")"
            Debug.Print "        m" & fld.Name & " = Valor"
            Debug.Print "    End Property"
            Debug.Print
        Next fld
    
        Debug.Print
        'Debug.Print "' ===================================================================="
        Debug.Print "End Class"
        Debug.Print "' ===================================================================="

    End If
    MsgBox "Clase generada en la ventana inmediata.", vbInformation

End Sub

Private Function TipoAccessAVBA(Tipo As dao.DataTypeEnum) As String

    Select Case Tipo
        Case dbText, dbMemo
            TipoAccessAVBA = "String"

        Case dbByte
            TipoAccessAVBA = "Byte"

        Case dbInteger
            TipoAccessAVBA = "Integer"

        Case dbLong
            TipoAccessAVBA = "Long"

        Case dbSingle
            TipoAccessAVBA = "Single"

        Case dbDouble
            TipoAccessAVBA = "Double"

        Case dbCurrency
            TipoAccessAVBA = "Currency"

        Case dbDate
            TipoAccessAVBA = "Date"

        Case dbBoolean
            TipoAccessAVBA = "Boolean"

        Case dbDecimal
            TipoAccessAVBA = "Variant" ' Access no expone Decimal directamente

        Case Else
            TipoAccessAVBA = "Variant"
    End Select

End Function

