VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   5
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        espera 0.5
        Me.Refresh
        
        'Vemos datos de configconta.ini
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then

             MsgBox "MAL CONFIGURADO", vbCritical
             End
             Exit Sub
        End If
        
         
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
         If AbrirConexionUsuarios() = False Then
             MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
             End
         End If
         
         'La llave
'         If True Then
'                Load frmLLave
'                If Not frmLLave.ActiveLock1.RegisteredUser Then
'                    'No ESTA REGISTRADO
'                    frmLLave.Show vbModal
'                Else
'                    Unload frmLLave
'                End If
'         End If
         
         
         
         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then espera T1

         
         PonerVisible True
         If Text1(0).Text <> "" Then
            Text1(1).SetFocus
        Else
            Text1(0).SetFocus
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PonerVisible False
    T1 = Timer
    Text1(0).Text = ""
    Text1(1).Text = ""
    Me.Label1(3).Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
    PrimeraVez = True
    CargaImagen

End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\arifonT.dat")
    Me.Height = Me.Image1.Height
    Me.Width = Me.Image1.Width
    
    'LOs text
    FijarText
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set Conn = Nothing
        End
    End If
End Sub

Private Sub FijarText()
Dim L As Long
    On Error GoTo EF
    L = Me.Width - Text1(0).Width - 120
    Text1(0).Left = L
    Text1(1).Left = L
    Me.Label1(0).Left = L
    Me.Label1(1).Left = L
    Me.Label1(2).Left = L
    
    
    L = Me.Height - Label1(2).Height - 120
    Me.Label1(2).Top = L
    Text1(1).Top = L
    L = L - 320   '375 + algo
    Label1(1).Top = L
    L = L - 350 '330+20
    Text1(0).Top = L
    L = L - 380
    Label1(0).Top = L
    
EF:
    If Err.Number <> 0 Then MuestraError Err.Number
        
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub









Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    End If
        
    
End Sub



Private Sub Validar()
Dim NuevoUsu As Usuario
Dim Ok As Byte

    'Validaremos el usuario y despues el password
    Set vUsu = New Usuario
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vUsu.PasswdPROPIO = Text1(1).Text Then
            Ok = 0
        Else
            Ok = 1
        End If

    Else
        Ok = 2
    End If
    
    If Ok <> 0 Then
        MsgBox "Usuario-Clave Incorrecto", vbExclamation

            Text1(1).Text = ""
            Text1(0).SetFocus
    Else
        'OK
        Screen.MousePointer = vbHourglass
        CadenaDesdeOtroForm = "OK"
        Label1(2).Caption = ""  'Si tarda pondremos texto aquin
        PonerVisible False
        Me.Refresh
        Screen.MousePointer = vbHourglass
        HacerAccionesBD
        Unload Me
    End If

End Sub

Private Sub HacerAccionesBD()
Dim SQL As String


    
''''    'Limpiamos datos blanace
''''    SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''    Me.Refresh
''''
''''    SQL = "DELETE from Usuarios.ztmpconextcab where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''    Me.Refresh
''''
''''    SQL = "DELETE from usuarios.ztmpconext where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''    Me.Refresh
''''
''''    SQL = "DELETE from Usuarios.zcuentas where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''
''''    Me.Refresh
''''    SQL = "DELETE from usuarios.ztmplibrodiario where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
    
    
End Sub


Private Sub PonerVisible(Visible As Boolean)
    Label1(2).Visible = Not Visible  'Cargando
    Text1(0).Visible = Visible
    Text1(1).Visible = Visible
    Label1(0).Visible = Visible
    Label1(1).Visible = Visible
End Sub




'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim Cad As String
On Error GoTo ENumeroEmpresaMemorizar


        
    Cad = App.Path & "\ultusuT.dat"
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = Cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = Text1(0).Text
        Print #NF, Cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub
