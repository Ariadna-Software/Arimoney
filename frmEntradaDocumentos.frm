VERSION 5.00
Begin VB.Form frmEntradaDocumentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción documentos"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   Icon            =   "frmEntradaDocumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optPagare 
      Caption         =   "Pagaré"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton optPagare 
      Caption         =   "Talón"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox txtDocumento 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtDocumento 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtDocumento 
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtDocumento 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Banco"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Importe"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Nº documento"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha recepcion"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Datos documento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   12
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2490
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   1560
      Top             =   1440
      Width           =   240
   End
End
Attribute VB_Name = "frmEntradaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
Dim SQL As String
    SQL = ""
    If Me.txtDocumento(3).Text = "" Then SQL = "Cliente" & vbCrLf
    If Me.txtVto(0).Text = "" Then SQL = SQL & "Numero serie" & vbCrLf
    If Me.txtVto(1).Text = "" Then SQL = SQL & "Numero factura" & vbCrLf
    
    If SQL <> "" Then
        MsgBox "Faltan campos: " & vbCrLf & SQL, vbExclamation
        Exit Sub
    End If
    
    'Ha puesto los datos
    Set miRsAux = New ADODB.Recordset
    SQL = "from scobro,sforpa where scobro.codforpa=sforpa.codforpa"
    SQL = SQL & " AND codmacta ='" & txtDocumento(3).Text & "'"
      
    'Y el tipo de pago
    


    
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Image2, 2
    Limpiar Me
End Sub




Private Sub optPagare_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtDocumento_GotFocus(Index As Integer)
    ObtenerFocoGral txtDocumento(Index)
End Sub

Private Sub txtDocumento_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyPressGral KeyAscii
End Sub

Private Sub txtDocumento_LostFocus(Index As Integer)
    txtDocumento(Index).Text = Trim(txtDocumento(Index).Text)
    If txtDocumento(Index).Text = "" Then
        'Codmacta
        If Index = 3 Then Me.txtAux(3).Text = ""
        
        Exit Sub  'Salimos
    End If
    
    
    Select Case Index
    Case 0, 1
        'Numero y banco
        'lo pasamos a mayusculas
        txtDocumento(Index).Text = UCase(txtDocumento(Index).Text)
            
    Case 2
        'importe
        FormatTextImporte txtDocumento(Index)
    
    Case 3
        'Codmacta
        If Not CuentaCorrectaUltimoNivelTXT(txtDocumento(Index), txtAux(3)) Then
            
            MsgBox txtAux(3).Text, vbExclamation
            txtDocumento(Index).Text = ""
            txtAux(3).Text = ""
            PonerFocoGral txtDocumento(Index)
        End If
    End Select
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    ObtenerFocoGral txtDocumento(Index)
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    If txtFecha(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
        txtFecha(Index).Text = ""
        Ponerfoco txtFecha(Index)
    End If
End Sub

