Attribute VB_Name = "__Motor_Test"
Option Compare Database
Option Explicit



' Declaraciones para compatibilidad con 32 y 64 bits
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Sub MedirTiempoPreciso()
    Dim Tstart As Currency, Tend As Currency, Tfreq As Currency
    Dim TiempoTotal As Double

    With CFG
        .ModoLigadura = 2
        .ModoH = 0
        .ModoLateral = 0
        .ModoSibilantes = 0
        .ModoX = 0
    End With

    ' Activar Logs
    DebugMotor = False
    DebugDTO = True
'    PreferirIlLu = True
    
    InitLog

    CargaCacheIPA
    
    ' Obtener la frecuencia del sistema (conteo por segundo)
    QueryPerformanceFrequency Tfreq
    
    ' Inicio del cronómetro
    QueryPerformanceCounter Tstart
    
    ' --- Coloca aquí el código que quieres medir ---
    'Entrada_Motor_ES ("El exalumno uruguayo oyó a Héctor y a Xavier hablar sobre la increíble hazaña de aquel héroe que, sin embargo, aún rehusaba ir a la ópera, aunque ella insistía en que era útil; mientras tanto, mi amigo Andrés, muy ilusionado, le explicó cómo se oye el ruido extraño del viejo xilófono azul.")
    
'    'SilabearFrase ("Pablo Diego José Francisco de Paula Juan Nepomuceno María de los Remedios Cipriano de la Santísima Trinidad Ruiz y Picasso")
    'Entrada_Motor_ES ("Pablo Diego José Francisco de Paula Juan Nepomuceno María de los Remedios Cipriano de la Santísima Trinidad Ruiz y Picasso")
    'Entrada_Motor_ES ("Triana, Jiménez Paitovi")
    'Entrada_Motor_ES ("José María de las Nieves Álvarez-Sotomayor y Villafranca del Río")
    'Entrada_Motor_ES ("Aurelio Ignacio Ezequiel de la Santísima Concepción del Sagrado Corazón de Jesús Fernández-Montemayor y Villalobos-Quintanilla de los Ríos")
    
    'Entrada_Motor_CA ("Maria Àngela Valentina Eulàlia Güell-i-Ferreró-Montjuïc")
    'Entrada_Motor_CA ("Eulàlia Maria Raül Saüdaqüerreixent")
    'Entrada_Motor_CA ("Joana Eloïsa Païsatge Heroïcitat Coïssor") ' como él no es de Ana su hijo mi amigo pero eso lo oigo muy alto se oye")
    'Entrada_Motor_CA ("Berenguera d’Òdena de Montserrat i Llorençà")
    'Entrada_Motor_CA ("Aüerònia d’Il·luïssà de Montsant-i-Corominesçó d’Òrrius-Joïaquim")
    'Entrada_Motor_IB ("Alba Salvá i Ripoll")
    'Entrada_Motor_IB ("Aina Margarida d’Escorca i Alcover-d’Es Puig de Son Garrit")
    'Entrada_Motor_VA ("Ausiàs Benimacletixotxa Almenarbrell")
    
    'Entrada_Motor_EU ("Etxeberriazarragaetxebarriandikoetxea elektroentzefalografistarenak autogobernuaren erdibideko azpiegiturak")
    'Entrada_Motor_EU ("Osabaetxeberriazarragaetxebarriandikoetxeazarretxandiarena")
    'Entrada_Motor_EU ("aurreantimikroelektroentzefalografistarenekoa")
    'Entrada_Motor_EU ("aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketautogobernuarenberregituraketaren")
    'Entrada_Motor_EU ("aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketautogobernuarenberregituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekin")
    'Entrada_Motor_EU ("Aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekinberregituraketaren")
    'Entrada_Motor_EU ("aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekinberregituraketarenazpimikroentzefalogramaren")
    'Entrada_Motor_EU ("Osabaetxeberriazarragaetxebarriandikoetxeazarretxandiarena Aurreantimikroelektroentzefalografistarenekoa Aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketautogobernuarenberregituraketaren Aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketautogobernuarenberregituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekin Aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekinberregituraketaren Aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekinberregituraketaren " & _
        " Aurreantimikroelektroentzefalogramagileenganakoetxeberriazarragaetxebarriandikoetxeetakoenpresaburuenartekoazpiegituraketarenazpimikroentzefalografistarenekogainelektromagnetikoarekinberregituraketarenazpimikroentzefalogramaren")
    
    'Entrada_Motor_EU ("Aitor Xabier Etxeberria------Goikoetxeandia Aranzabalbeitia ")
    'Entrada_Motor_EU ("Iñigo Joseba Etxeandia-Larrinagaberrizabal Odriozolakortabarria")
    'Entrada_Motor_EU ("Miren Ane Urrutikoetxea-Zumarragagoikoetxeaga Mendizabalbeitia")
    'Entrada_Motor_EU ("Jon Ander Agirregomezkorta-Lizarraldebeitia Etxeberriandikoetxea")
    'Entrada_Motor_EU ("Ane Miren Goizargi Arrieta-Goikoetxeandikoetxea Urrutikoetxeaga")
    
    'Entrada_Motor_GL ("Xoán Uxía Anxo Xiana Xurxo Noaia")
    'Entrada_Motor_GL ("Xoío Aiaía Seixoiro Lingüística Pingüía Anhxoxo Llxexo Xrroxo Fêlix")
    'Entrada_Motor_GL ("Ameixeiras Xunqueira Meixide Queiruga Oubiña Figueroa Carbalho Ferreiro Carballeira")
    'Entrada_Motor_GL ("Oubiña Carbalho")
    'Entrada_Motor_GL ("")
    'Entrada_Motor_GL ("")
    'Entrada_Motor_GL ("")
    'Entrada_Motor_GL ("")
    'Entrada_Motor_GL ("")
    'Entrada_Motor_GL ("")
    
' -----------------------------------------------
    ' Fin del cronómetro
    QueryPerformanceCounter Tend
    
    ' Cálculo del tiempo en segundos (con decimales de microsegundos)
    TiempoTotal = (Tend - Tstart) / Tfreq
    
    ObjDTO.Tiempo = CStr(TiempoTotal)
    addLog
    addLog "Tiempo transcurrido: " & Format(TiempoTotal, "0.000000") & " segundos"
    
    PrintLog
    'Debug.Print
    Debug.Print "Tiempo transcurrido: " & Format(TiempoTotal, "0.000000") & " segundos"
    
    DoCmd.OpenForm "frmSalida", acNormal, , , , acHidden
    
    With Form_frmSalida
        .lblOriginal.Caption = ObjDTO.TextoOriginal
        .txtSilabas = ObjDTO.SilabasFinal
        .txtFonemas = ObjDTO.FonemasFinal
        .lblTiempo.Caption = "Tiempo proceso: " & Format(ObjDTO.Tiempo, "0.000000") & " segundos"
        .cmdSalir.SetFocus
        .Visible = True
    End With
    
    If DebugDTO Then
        Call MF_DebugDTO("Motor GL")
    End If
    
End Sub




