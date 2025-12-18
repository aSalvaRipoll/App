Attribute VB_Name = "21_modReferencias"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 21_modReferencias
' Gestión completa de referencias del proyecto VBA
'===============================================================
' Incluye:
'   - Listado de referencias (VBIDE y Access)
'   - Comprobación por ruta o GUID
'   - Agregar referencia por ruta
'   - Agregar referencia por GUID
'   - Asegurar referencia (si no existe, agregarla)
'   - Borrar referencia por GUID
'   - Detección de referencias rotas
'===============================================================


'---------------------------------------------------------------
' LISTAR REFERENCIAS (VBIDE)
'---------------------------------------------------------------
Public Sub ListarReferenciasVBIDE(frm As Form)
    Dim objRef As Reference
    Dim objRefs As References

    frm.lstRefs.RowSource = ""
    frm.lstRefs.Requery

    Set objRefs = Application.VBE.ActiveVBProject.References

    For Each objRef In objRefs
        frm.lstRefs.AddItem _
            objRef.Name & " | " & _
            objRef.Description & " | " & _
            IIf(objRef.BuiltIn, "Interna", "Externa") & " | " & _
            objRef.Guid & " | " & _
            IIf(objRef.IsBroken, "ROTA", objRef.FullPath)
    Next objRef

    Set objRefs = Nothing
End Sub


'---------------------------------------------------------------
' LISTAR REFERENCIAS (Access)
'---------------------------------------------------------------
Public Sub ListarReferenciasAccess(frm As Form)
    Dim objRef As Reference

    frm.lstRefs.RowSource = ""
    frm.lstRefs.Requery

    For Each objRef In References
        frm.lstRefs.AddItem _
            objRef.Name & " | " & _
            objRef.FullPath & " | " & _
            objRef.Guid & " | " & _
            IIf(objRef.IsBroken, "ROTA", "OK")
    Next objRef
End Sub


'---------------------------------------------------------------
' COMPROBAR REFERENCIA POR RUTA O GUID
'---------------------------------------------------------------
Public Function ComprobarReferencias(Optional strAdd As String, Optional strGUID As String) As Boolean
    Dim objRef As Reference
    Dim objRefs As References

    ComprobarReferencias = False

    If strAdd = "" And strGUID = "" Then
        MsgBox "Debe indicar al menos un parámetro (ruta o GUID).", vbExclamation
        Exit Function
    End If

    Set objRefs = Application.VBE.ActiveVBProject.References

    For Each objRef In objRefs

        If strAdd <> "" Then
            If StrComp(objRef.FullPath, strAdd, vbTextCompare) = 0 Then
                ComprobarReferencias = True
                Exit Function
            End If
        End If

        If strGUID <> "" Then
            If UCase$(objRef.Guid) = UCase$(strGUID) Then
                ComprobarReferencias = True
                Exit Function
            End If
        End If

    Next objRef

    Set objRefs = Nothing
End Function


'---------------------------------------------------------------
' AGREGAR REFERENCIA POR RUTA
'---------------------------------------------------------------
Public Function AgregarReferenciaPorRuta(ruta As String) As Boolean
    Dim vbProj As VBIDE.VBProject

    On Error GoTo ErrHandler

    Set vbProj = Application.VBE.ActiveVBProject
    vbProj.References.AddFromFile ruta

    Debug.Print ">> Referencia agregada: "; ruta
    AgregarReferenciaPorRuta = True
    Exit Function

ErrHandler:
    Debug.Print ">> No se pudo agregar la referencia: "; ruta; " - "; Err.Description
    AgregarReferenciaPorRuta = False
End Function


'---------------------------------------------------------------
' AGREGAR REFERENCIA POR GUID
'---------------------------------------------------------------
Public Function AgregarReferenciaPorGUID(strGUID As String, _
                                         Optional major As Long = 1, _
                                         Optional minor As Long = 0) As Boolean
    Dim vbProj As VBIDE.VBProject

    On Error GoTo ErrHandler

    Set vbProj = Application.VBE.ActiveVBProject
    vbProj.References.AddFromGuid strGUID, major, minor

    Debug.Print ">> Referencia agregada por GUID: "; strGUID
    AgregarReferenciaPorGUID = True
    Exit Function

ErrHandler:
    Debug.Print ">> No se pudo agregar la referencia por GUID: "; strGUID; " - "; Err.Description
    AgregarReferenciaPorGUID = False
End Function


'---------------------------------------------------------------
' ASEGURAR REFERENCIA (si no existe, agregarla)
'---------------------------------------------------------------
Public Function AsegurarReferenciaPorGUID(strGUID As String, _
                                          Optional major As Long = 1, _
                                          Optional minor As Long = 0) As Boolean

    If ComprobarReferencias(, strGUID) Then
        AsegurarReferenciaPorGUID = True
    Else
        AsegurarReferenciaPorGUID = AgregarReferenciaPorGUID(strGUID, major, minor)
    End If
End Function


'---------------------------------------------------------------
' BORRAR REFERENCIA POR GUID
'---------------------------------------------------------------
Public Function BorrarReferencia(strGUID As String) As Boolean
    Dim objRef As Reference
    Dim vbProj As VBIDE.VBProject

    On Error GoTo ErrHandler

    Set vbProj = Application.VBE.ActiveVBProject

    For Each objRef In vbProj.References
        If UCase$(objRef.Guid) = UCase$(strGUID) Then
            vbProj.References.Remove objRef
            Debug.Print ">> Referencia eliminada: "; objRef.Name
            BorrarReferencia = True
            Exit Function
        End If
    Next objRef

    Debug.Print ">> No se encontró la referencia con GUID: "; strGUID
    Exit Function

ErrHandler:
    Debug.Print ">> No se pudo eliminar la referencia: "; Err.Description
    BorrarReferencia = False
End Function


'---------------------------------------------------------------
' DETECTAR REFERENCIAS ROTAS
'---------------------------------------------------------------
Public Function HayReferenciasRotas() As Boolean
    Dim objRef As Reference

    For Each objRef In Application.VBE.ActiveVBProject.References
        If objRef.IsBroken Then
            HayReferenciasRotas = True
            Exit Function
        End If
    Next objRef
End Function


