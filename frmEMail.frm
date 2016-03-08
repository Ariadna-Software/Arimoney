VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar E-MAIL"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2940
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5715
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   2055
         Index           =   3
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Para"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   600
         Picture         =   "frmEMail.frx":000C
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   5715
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Otro"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Error"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sugerencia"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-Mail Ariadna Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmEMail.frx":685E
      Top             =   3840
      Width           =   480
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '0 - Envio del PDF
    '1 - Envio Mail desde menu soporte
    '2 -
    '3 - Reclamaciones via E-MAIL
        'Valores en CadenaDesdeOtroForm
Public MisDatos As String
    'Nombre para|email para|Asunto|Mensaje|
    
    
Public queEmpresa     As Byte
    '0  Cualquiera
    '1  Escalona
    
    
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Dim Cad As String
Dim PrimeraVez As Boolean
Dim DatosDelMailEnUsuario As String

Private Sub Enviar()
    Dim imageContentID, success
    Dim mailman As ChilkatMailMan
    Dim Valores As String
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante, es la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
    mailman.LogMailSentFilename = "" ' App.Path & "\mailSent.log"

    
    'Servidor smtp
    Valores = DatosDelMailEnUsuario  'Empipado: smtphost,smtpuser, pass, diremail
    If Valores = "" Then
        MsgBox "Falta configurar en paremtros la opcion de envio mail(servidor, usuario, clave)"
        Exit Sub
    End If
    
    'ANTES PRUEBAS GMAIL
    mailman.SmtpHost = RecuperaValor(Valores, 2) ' vParam.SmtpHOST
    mailman.SmtpUsername = RecuperaValor(Valores, 1) 'vParam.SmtpUser
    mailman.SmtpPassword = RecuperaValor(Valores, 3) 'vParam.SmtpPass
    'David 2 Mayo 2007
    mailman.SmtpAuthMethod = "LOGIN"
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    'Si es de SOPORTE
    If Opcion = 1 Then
         'Obtenemos la pagina web de los parametros
        Cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
        If Cad = "" Then
            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
            Exit Sub
        End If
    
        If Cad = "" Then GoTo GotException
        email.AddTo "Soporte Tesoreria", Cad
        Cad = "Soporte Ariconta. "
        If Option1(0).Value Then Cad = Cad & Option1(0).Caption
        If Option1(1).Value Then Cad = Cad & Option1(1).Caption
        If Option1(2).Value Then Cad = Cad & "Otro: " & Text2.Text
        email.Subject = Cad
        
        'Ahora en text1(3).text generaremos nuestro mensaje
        Cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
        Cad = Cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
        Cad = Cad & "TESORERIA:  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
        Cad = Cad & "Usuario: " & vUsu.Nombre & vbCrLf
        Cad = Cad & "Nivel USU: " & vUsu.Nivel & vbCrLf
        Cad = Cad & "Empresa: " & vEmpresa.nomempre & vbCrLf
        Cad = Cad & "&nbsp;<hr>"
        Cad = Cad & Text3.Text & vbCrLf & vbCrLf
        Text1(3).Text = Cad
    Else
        'Envio de mensajes normal
        email.AddTo Text1(0).Text, Text1(1).Text
        Cad = Text1(2).Text
        If queEmpresa = 1 Then Cad = Cad & " [ARI]"
        email.Subject = Cad
    End If
    

    
   
    
    'El resto lo hacemos comun
    'La imagen
    'imageContentID = email.AddRelatedContent(App.Path & "\minilogo.bmp")
        
        
    'Comun
    Cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    Cad = Cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P><FONT FACE=""Tahoma""><FONT SIZE=3>"
    FijarTextoMensaje
    Cad = Cad & "</FONT></FONT></P></TD></TR><TR><TD VALIGN=""TOP"">"

    
    
    
    Select Case queEmpresa
    Case 1
        
        Cad = Cad & "<p class=""MsoNormal""><b><i>"
        Cad = Cad & "<span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">C."
        Cad = Cad & "R. Reial Séquia Escalona</span></i></b></p>"
        Cad = Cad & "<p class=""MsoNormal""><em><b>"
        Cad = Cad & "<span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">"
        Cad = Cad & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; La Junta</span></b></em><span style=""font-size: 10.0pt; font-family: Arial,sans-serif; color: black"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">&nbsp;</span></p>"
        Cad = Cad & "<p class=""MsoNormal"">"
        Cad = Cad & "<span style=""font-size: 13.5pt; font-family: Arial,sans-serif; color: #9999FF"">"
        Cad = Cad & "********************</span></p>"
         Cad = Cad & "<p class=MsoNormal><b>"
         Cad = Cad & "<span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>Confidencialidad"
         Cad = Cad & "</span></b><span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'><br>"
         Cad = Cad & "Este mensaje y sus archivos adjuntos van dirigidos exclusivamente a su destinatario, "
         Cad = Cad & "pudiendo contener información confidencial sometida a secreto profesional. No está permitida su reproducción o "
         Cad = Cad & "distribución sin la autorización expresa de Real Acequia Escalona. Si usted no es el destinatario final por favor "
         Cad = Cad & "elimínelo e infórmenos por esta vía.<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-top-alt:6.0pt;"
         Cad = Cad & "margin-right:0cm;margin-bottom:6.0pt;margin-left:0cm;text-align:justify'><span style='font-size:8.0pt;"
         Cad = Cad & "font-family:""Comic Sans MS"";color:black'>De acuerdo con la Ley 34/2002 (LSSI) y la Ley 15/1999 (LOPD), "
         Cad = Cad & "usted tiene derecho al acceso, rectificación y cancelación de sus datos personales informados en el fichero del que es "
         Cad = Cad & "titular Real Acequia Escalona. Si desea modificar sus datos o darse de baja en el sistema de comunicación electrónica "
         Cad = Cad & "envíe un correo a</span> <span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>"
         Cad = Cad & "<a href=""mailto:escalona@acequiaescalona.org"">escalona@acequiaescalona.org</a> </span><span style='font-size:8.0pt;"
         Cad = Cad & "font-family:""Comic Sans MS""'>, <span style='color:black'>indicando en la línea de <b>&#8220;Asunto&#8221;</b> el derecho "
         Cad = Cad & "que desea ejercitar. <o:p></o:p></span></span></p><p class=MsoNormal><o:p>&nbsp;</o> "
         
         'ahora en valenciano
         Cad = Cad & ""
         Cad = Cad & "<p class=MsoNormal><b>"
         Cad = Cad & "<span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>Confidencialitat"
         Cad = Cad & "</span></b><span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'><br>"
         Cad = Cad & "Aquest missatge i els seus arxius adjunts van dirigits exclusivamente al seu destinatari, "
         Cad = Cad & "podent contindre informació confidencial sotmesa a secret professional. No està permesa la seua reproducció o "
         Cad = Cad & "distribució sense la autorització expressa de Reial Séquia Escalona. Si vosté no és el destinatari final per favor "
         Cad = Cad & "elimíneu-lo e informe-nos per aquesta via.<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-top-alt:6.0pt;"
         Cad = Cad & "margin-right:0cm;margin-bottom:6.0pt;margin-left:0cm;text-align:justify'><span style='font-size:8.0pt;"
         Cad = Cad & "font-family:""Comic Sans MS"";color:black'>D'acord amb la Llei 34/2002 (LSSI) i la Llei 15/1999 (LOPD), "
         Cad = Cad & "vosté té dret a l'accés, rectificació i cancelació de les seues dades personals informats en el ficher del qué és "
         Cad = Cad & "titolar Reial Séquia Escalona. Si vol modificar les seues dades o donar-se de baixa en el sistema de comunicació electrònica "
         Cad = Cad & "envíe un correu a</span> <span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>"
         Cad = Cad & "<a href=""mailto:escalona@acequiaescalona.org"">escalona@acequiaescalona.org</a> </span><span style='font-size:8.0pt;"
         Cad = Cad & "font-family:""Comic Sans MS""'>, <span style='color:black'>indicant en la línea de <b>&#8220;Asumpte&#8221;</b> el dret "
         Cad = Cad & "que desitja exercitar. <o:p></o:p></span></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p> "
         

        

    
    
    Case Else
        
        
        Cad = Cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
        
        Cad = Cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa " & App.EXEName & " de "
        Cad = Cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
        Cad = Cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
        
        Cad = Cad & "<P>Este correo electrónico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
        Cad = Cad & " los destinatarios especificados. La información contenida puesde ser CONFIDENCIAL"
        Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
        Cad = Cad & "<P>Si usted recibe este mensaje por ERROR, por favor comuníqueselo inmediatamente al"
        Cad = Cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelación, distribución"
        Cad = Cad & " impresión o copia de toda o alguna parte de la información contenida, Gracias "
        Cad = Cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
        
        
        Cad = Cad & "<HR ALIGN=""LEFT"" SIZE=1></TD>"
        
            
            
        
    
    End Select
    'Fianl
    Cad = Cad & "</TR></BODY></HTML>"
    
    email.SetHtmlBody (Cad)
    
    
    
    'Texto alternativo
    Select Case queEmpresa
    Case 1
    
    
    Case Else
        
    End Select
    'Texto alternativo
    Cad = ""
    Cad = Cad & "Este correo electronico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a " & vbCrLf
    Cad = Cad & " los destinatarios especificados. La informacion contenida puesde ser CONFIDENCIAL" & vbCrLf
    Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA." & vbCrLf & vbCrLf
    Cad = Cad & "Si usted recibe este mensaje por ERROR, por favor comuniqueselo inmediatamente al" & vbCrLf
    Cad = Cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelacion, distribucion" & vbCrLf
    Cad = Cad & " impresion o copia de toda o alguna parte de la informacion contenida, Gracias " & vbCrLf
    
    email.AddPlainTextAlternativeBody Text1(3).Text & vbCrLf & vbCrLf & vbCrLf & Cad
    email.From = RecuperaValor(Valores, 1) 'vParam.diremail
    
    If Opcion = 0 Or Opcion = 3 Then
        'ADjunatmos el PDF
        email.AddFileAttachment App.Path & "\docum.pdf"
    End If
        
  
   

    If vParam.EnvioDesdeOutlook Then
        
         mailman.SendViaOutlook email
         success = 1
    Else
        'Esta es la que estaba
        success = mailman.SendEmail(email)
        
    End If
    If (success = 1) Then
        If Opcion <> 3 Then
            If vParam.EnvioDesdeOutlook Then
                Cad = "Enviado al outlook"
            Else
                Cad = "Mensaje enviado correctamente."
            End If
            MsgBox Cad, vbInformation
            Command2(0).SetFocus
        End If
    Else
        Cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.Path & "\log.xml"
        MsgBox Cad, vbExclamation
    End If
    
    
GotException:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        NumRegElim = 1
    Else
        NumRegElim = 0
    End If
    Set email = Nothing
    Set mailman = Nothing

End Sub

Private Sub Command1_Click()
    If Not DatosOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    Image2.Visible = True
    Me.Refresh
    Enviar
    Image2.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 3 Then
            HacerMultiEnvio
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Image2.Visible = False
    Limpiar Me
    Frame1(0).Visible = (Opcion = 0) Or (Opcion = 3)
    Frame1(1).Visible = (Opcion = 1)
    If Opcion = 1 Then HabilitarText
    Me.Icon = frmPpal.Icon
    PonDisponibilidadEmail
    Me.Command1.Enabled = (DatosDelMailEnUsuario <> "")
    If Opcion = 3 Then
        'Si es masivo na de na
        Command2(0).Enabled = False
        Me.Command1.Enabled = False
        
    End If
End Sub




Private Sub PonDisponibilidadEmail()
    
   If vParam.EnvioDesdeOutlook Then
        Cad = "||||"
    Else
        Cad = DevuelveDesdeBD("dirfich", "Usuarios.usuarios", "codusu", (vUsu.Codigo Mod 100), "N")
        If Cad = "" Then
            'Primero compruebo si los datos los tengo en el usuario
            Cad = "select diremail,smtphost,smtpuser,smtppass from parametros"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux!SmtpHost) Then
                    For NumRegElim = 0 To miRsAux.Fields.Count - 1
                        Cad = Cad & DBLet(miRsAux.Fields(NumRegElim), "T") & "|"
                    Next NumRegElim
                End If
            End If
            miRsAux.Close
            Set miRsAux = Nothing
            
        End If
    
    End If
    DatosDelMailEnUsuario = Cad
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
    queEmpresa = 0
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)

    Screen.MousePointer = vbHourglass
    Text1(0).Tag = RecuperaValor(CadenaSeleccion, 1)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 2)
    'Si regresa con datos tengo k devolveer desde la bd el campo e-mail
    Cad = DevuelveDesdeBD("maidatos", "cuentas", "codmacta", Text1(0).Tag)
    Text1(1).Text = Cad
    Screen.MousePointer = vbDefault
End Sub

Private Sub Image1_Click()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 5  'NUEVO opcion
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(0).Text <> "" Then Text1(2).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
    HabilitarText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Function DatosOk() As Boolean
Dim I As Integer

    DatosOk = False
    If Opcion = 0 Then
                'Pocas cosas a comprobar
                For I = 0 To 2
                    Text1(I).Text = Trim(Text1(I).Text)
                    If Text1(I).Text = "" Then
                        MsgBox "El campo: " & Label1(I).Caption & " no puede estar vacio.", vbExclamation
                        Exit Function
                    End If
                Next I
                
                'EL del mail tiene k tener la arroba @
                I = InStr(1, Text1(1).Text, "@")
                If I = 0 Then
                    MsgBox "Direccion e-mail erronea", vbExclamation
                    Exit Function
                End If
    Else
        Text2.Text = Trim(Text2.Text)
        'SOPORTE
        If Trim(Text3.Text) = "" Then
            MsgBox "El mensaje no puede ir en blanco", vbExclamation
            Exit Function
        End If
        If Option1(2).Value Then
            If Text2.Text = "" Then
                MsgBox "El campo 'OTRO asunto' no puede ir en blanco", vbExclamation
                Exit Function
            End If
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then Exit Sub 'Si estamos en el de datos nos salimos
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'El procedimiento servira para ir buscando los vbcrlf y cambiarlos por </p><p>
Private Sub FijarTextoMensaje()
Dim I As Integer
Dim J As Integer

    J = 1
    Do
        I = InStr(J, Text1(3).Text, vbCrLf)
        If I > 0 Then
              Cad = Cad & Mid(Text1(3).Text, J, I - J) & "</P><P>"
        Else
            Cad = Cad & Mid(Text1(3).Text, J)
        End If
        J = I + 2
    Loop Until I = 0
End Sub

Private Sub HabilitarText()
    If Option1(2).Value Then
        Text2.Enabled = True
        Text2.BackColor = vbWhite
    Else
        Text2.Enabled = False
        Text2.BackColor = &H80000018
    End If
End Sub



'Private Function RecuperarDatosEMAILAriadna() As Boolean
'Dim NF As Integer
'
'    RecuperarDatosEMAILAriadna = False
'    NF = FreeFile
'    Open App.Path & "\soporte.dat" For Input As #NF
'    Line Input #NF, cad
'    Close #NF
'    If cad <> "" Then RecuperarDatosEMAILAriadna = True
'
'End Function


'Private Function ObtenerValoresEnvioMail() As String
'    ObtenerValoresEnvioMail = ""
'    Set miRsAux = New ADODB.Recordset
'    cad = "Select diremail,SmtpHost, SmtpUser, SmtpPass  from parametros where"
'    cad = cad & " fechaini='" & Format(vParam.fechaini, FormatoFecha) & "';"
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not miRsAux.EOF Then
'        cad = DBLet(miRsAux!SmtpHost)
'        cad = cad & "|" & DBLet(miRsAux!SmtpUser)
'        cad = cad & "|" & DBLet(miRsAux!SmtpPass)
'        cad = cad & "|" & DBLet(miRsAux!diremail) & "|"
'        ObtenerValoresEnvioMail = cad
'    End If
'    miRsAux.Close
'    Set miRsAux = Nothing
'End Function

Private Sub HacerMultiEnvio()
Dim RS As ADODB.Recordset
Dim CONT As Integer
Dim I As Integer
    Cad = "Select tmp347.*,razosoci,maidatos from tmp347,cuentas where codusu =" & vUsu.Codigo
    Cad = Cad & " AND cuentas.codmacta = tmp347.cta AND (Importe is null)"
    
    'Campos comunes
    'ENVIO MASIVO DE EMAILS
    Text1(2).Text = RecuperaValor(MisDatos, 1)
    Text1(3).Text = RecuperaValor(MisDatos, 2)
    
    Me.Refresh
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText

    CONT = 0
    While Not RS.EOF
        CONT = CONT + 1
        RS.MoveNext
    Wend
    RS.MoveFirst
    I = 1
    Me.Refresh
    While Not RS.EOF
        Screen.MousePointer = vbHourglass
        Text1(0).Text = RS!razosoci
        Text1(1).Text = RS!maidatos
        Caption = "Enviar E-MAIL (" & I & " de " & CONT & ")"
        Me.Refresh
        
        'De momento volvemos a copiar el archivo como docum.pdf
        If Dir(App.Path & "\docum.pdf") <> "" Then
            Kill App.Path & "\docum.pdf"
            espera 0.3
        End If
        FileCopy App.Path & "\temp\" & RS!NIF, App.Path & "\docum.pdf"
        Me.Refresh
        NumRegElim = 0
        Enviar
        
        
        If NumRegElim = 1 Then
            'NO SE HA ENVIADO.
            Cad = "UPDATE tmp347 SET IMporte=0 WHERE codusu =" & vUsu.Codigo & " AND cliprov =0 AND cta='" & RS!Cta & "'"
            Conn.Execute Cad
        End If
        'Siguiente
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
End Sub


