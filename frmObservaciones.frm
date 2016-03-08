VERSION 5.00
Begin VB.Form frmObservaciones 
   Caption         =   "Observaciones"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAccion 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   0
      Left            =   6360
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmObservaciones.frx":0000
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "frmObservaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAccion_Click(Index As Integer)
    If Index = 0 Then
        CadenaDesdeOtroForm = Text1.Text
    
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    Text1.Text = CadenaDesdeOtroForm
    Text1.SelStart = Len(Text1.Text)
End Sub
