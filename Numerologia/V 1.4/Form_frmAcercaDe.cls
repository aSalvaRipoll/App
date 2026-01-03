VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    Dim strNotas As String
    Dim Acercade As String
    
    Me.Caption = "Sistema de Numerología Fonética (1991–)"

    
    Acercade = "Sistema fonético y numerológico diseñado por Alba Salvá Ripoll." & vbCrLf & vbCrLf & _
               "Pensado para ser claro, reversible y honesto. " & vbCrLf & vbCrLf & _
               "Si algún día continúas este trabajo, hazlo a tu manera," & vbCrLf & _
               "y sigue mis pasos y mi legado al mundo."
               
    
    strNotas = "<ul>" & vbCrLf & _
                    "<li>Este sistema nació de mi amor por las lenguas, por la claridad" & vbCrLf & _
                        "y por la belleza de la estructura.</li>" & vbCrLf & _
                    "<li>Cada regla fonética está pensada para ser coherente, reversible" & vbCrLf & _
                        "y con <b>respeto</b> por cada dialecto y cada nombre.</li>" & vbCrLf & _
                    "<li>Prefiero un sistema honesto y transparente antes que uno" & vbCrLf & _
                        "<b>""mágico""</b> pero opaco. Aquí todo se puede seguir, revisar y entender.</li>" & vbCrLf & _
                    "<li>Las excepciones <i>no son errores:</i> son historias, personas, lugares" & vbCrLf & _
                        "y formas de hablar que merecen ser escuchadas.</li>" & vbCrLf & _
                    "<li>Si algún día cambias algo de lo que hice, hazlo con buena intención," & vbCrLf & _
                        "con criterio y sin miedo a poner tu propia voz.</li>" & vbCrLf & _
                    "<li>Si alguna vez dudas, busca la coherencia. Ella <u>nunca te traicionará.</u></li>" & vbLf & _
                    "<li>Este sistema es también un puente: entre lenguas, entre generaciones" & vbCrLf & _
                        "y entre formas de mirar el mundo.</li>" & vbCrLf & _
                    "<li>Gracias por abrirlo, leerlo y, quizá, continuarlo. Eso ya hace" & vbCrLf & _
                        "que todo este trabajo haya valido la pena.</li>" & vbCrLf & _
                "</ul>"

        
    
    Me.txLegado = Acercade



    Me.txtMensaje = strNotas '"<ul><li>Este sistema nació de mi amor por las lenguas y por la claridad.</li>" & vbLf & _
                    "<li>Cada regla fonética fue pensada con <b>respeto</b> por su dialecto.</li>" & vbLf & _
                    "<li>Si alguna vez dudas, busca la coherencia. Ella <u>nunca te traicionará.</u></li>" & vbLf & _
                    "<li>No tengas miedo de cambiar lo que hice. Pero hazlo con buenas intenciones.</li></ul>"


    Me.cmdCerrar.SetFocus
    

End Sub
