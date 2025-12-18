Attribute VB_Name = "BasReparaClases"
Option Compare Database

Public Sub ReconstruirClasesLimpias()
    Dim comp As Object
    Dim rutaTemp As String
    Dim rutaOriginal As String
    Dim rutaNueva As String
    Dim nombre As String
    
    rutaTemp = CurrentProject.Path & "\_clases_reconstruidas_limpias\"
    
    
    Debug.Print "Reconstruyendo clases en: " & rutaTemp
    
       
    ' 1. Eliminar clases antiguas del proyecto
    Debug.Print
    Debug.Print "Eliminando clases antiguas..."
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Type = vbext_ct_ClassModule Then
            Application.VBE.ActiveVBProject.VBComponents.Remove comp
        End If
    Next comp
    
    ' 2. Importar clases originales
    Debug.Print
    Debug.Print "Importando clases Originales..."
    Dim archivo As String
    
    archivo = Dir(rutaTemp & "*.cls")

    Do While archivo <> ""
        
        ' Saltar archivos originales
        If InStr(1, archivo, "_ORIGINAL", vbTextCompare) = 0 Then
            
            Debug.Print
            Debug.Print "Borrando: "; archivo
            If Len(Dir(rutaTemp & archivo)) > 0 Then Kill rutaTemp & archivo
            
        Else 'Renombrar e importar
            
            Debug.Print
            
            rutaOriginal = Replace(archivo, "_ORIGINAL", "")
            Debug.Print "Renombrando: "; archivo; " -> "; rutaOriginal
            
            Name rutaTemp & archivo As rutaTemp & rutaOriginal
            
            Debug.Print
            Debug.Print "Archivo renombrado: "; rutaOriginal
            
            Debug.Print
            Debug.Print "Importando: "; rutaOriginal
            Application.VBE.ActiveVBProject.VBComponents.Import rutaTemp & rutaOriginal
        
            
        End If
        
        archivo = Dir
    Loop

    MsgBox "Reconstrucción completada. Todas las clases han sido regeneradas correctamente y limpiamente.", vbInformation
End Sub
