VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobrosAgentesLin 
      Height          =   3255
      Left            =   120
      TabIndex        =   545
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton cmdCobrosAgenLin 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   550
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   37
         Left            =   3600
         TabIndex        =   547
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   36
         Left            =   1320
         TabIndex        =   546
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   555
         Text            =   "Text1"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   549
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   552
         Text            =   "Text1"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   548
         Text            =   "Text1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   45
         Left            =   4200
         TabIndex        =   551
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   120
         TabIndex        =   561
         Top             =   2760
         Width           =   2505
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cobros agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   24
         Left            =   360
         TabIndex        =   560
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   43
         Left            =   2760
         TabIndex        =   559
         Top             =   975
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   37
         Left            =   3360
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   42
         Left            =   480
         TabIndex        =   558
         Top             =   975
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   36
         Left            =   1080
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  cobro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   82
         Left            =   240
         TabIndex        =   557
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   480
         TabIndex        =   556
         Top             =   2085
         Width           =   465
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   7
         Left            =   1080
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   48
         Left            =   480
         TabIndex        =   554
         Top             =   1725
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   80
         Left            =   480
         TabIndex        =   553
         Top             =   1440
         Width           =   615
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   6
         Left            =   1080
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Frame FramereclaMail 
      Height          =   6735
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   10755
      Begin VB.CheckBox chkExcluirConEmail 
         Caption         =   "Excluir clientes con email (carta)"
         Height          =   255
         Left            =   7560
         TabIndex        =   110
         Top             =   5520
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   83
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   82
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkReclamaDevueltos 
         Caption         =   "Solo devueltos"
         Height          =   255
         Left            =   6600
         TabIndex        =   105
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   94
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   95
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   96
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   3
         Left            =   4320
         TabIndex        =   97
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   4
         Left            =   7320
         TabIndex        =   100
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   5
         Left            =   6240
         TabIndex        =   99
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   6
         Left            =   5280
         TabIndex        =   98
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   7
         Left            =   8400
         TabIndex        =   101
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkMostrarCta 
         Caption         =   "Mostrar cuenta"
         Height          =   255
         Left            =   8520
         TabIndex        =   106
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CheckBox chkInsertarReclamas 
         Caption         =   "Insertar registros reclamaciones"
         Height          =   195
         Left            =   4800
         TabIndex        =   109
         Top             =   5550
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.TextBox txtVarios 
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   6240
         Width           =   2775
      End
      Begin VB.TextBox txtVarios 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   6240
         Width           =   4215
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Enviar por e-mail"
         Height          =   255
         Left            =   3000
         TabIndex        =   108
         Top             =   5520
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkMarcarUtlRecla 
         Caption         =   "Marcar fecha ultima reclamacion"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   5520
         Width           =   2655
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Left            =   2040
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox txtDescCarta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   139
         Text            =   "Text2"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtCarta 
         Height          =   285
         Left            =   2880
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   135
         Text            =   "Text1"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   7320
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   10
         Left            =   9600
         TabIndex        =   87
         Text            =   "99/99/9999"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   7680
         TabIndex        =   86
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   600
         TabIndex        =   102
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   9480
         TabIndex        =   114
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdreclama 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   8280
         TabIndex        =   113
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   5
         Left            =   6480
         TabIndex        =   89
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   5400
         TabIndex        =   85
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   88
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   3600
         TabIndex        =   84
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7560
         TabIndex        =   118
         Text            =   "Text5"
         Top             =   1860
         Width           =   3075
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   117
         Text            =   "Text5"
         Top             =   1920
         Width           =   3075
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   116
         Text            =   "Text1"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   7320
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   41
         Left            =   1440
         TabIndex        =   544
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   40
         Left            =   240
         TabIndex        =   543
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   10
         Left            =   9360
         Top             =   1102
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   79
         Left            =   240
         TabIndex        =   542
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   18
         Left            =   4680
         TabIndex        =   142
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   141
         Top             =   6000
         Width           =   600
      End
      Begin VB.Image imgCarta 
         Height          =   240
         Left            =   3480
         Picture         =   "frmListado.frx":000C
         ToolTipText     =   "Seleccionar tipo carta"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   16
         Left            =   2160
         TabIndex        =   140
         Top             =   4680
         Width           =   360
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   2
         Left            =   6120
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   5
         Left            =   6120
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   4
         Left            =   840
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   7
         Left            =   5160
         Top             =   1102
         Width           =   240
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   138
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   137
         Top             =   2805
         Width           =   465
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   2
         Left            =   6120
         Picture         =   "frmListado.frx":685E
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   5520
         TabIndex        =   136
         Top             =   2805
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   8880
         TabIndex        =   133
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vencimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   14
         Left            =   6720
         TabIndex        =   132
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   131
         Top             =   1095
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   9
         Left            =   7440
         Top             =   1102
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Reclamaci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   130
         Top             =   4680
         Width           =   1620
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   8
         Left            =   240
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   129
         Top             =   1965
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   5520
         TabIndex        =   127
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   126
         Top             =   1095
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   6
         Left            =   3360
         Top             =   1102
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "R E C L A M A C I O N E S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   2
         Left            =   2640
         TabIndex        =   125
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   124
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   11
         Left            =   2760
         TabIndex        =   123
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   122
         Top             =   3360
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   121
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   120
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   2880
         TabIndex        =   119
         Top             =   4680
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   128
         Top             =   1095
         Width           =   615
      End
   End
   Begin VB.Frame FrCobrosPendientesCli 
      Height          =   7455
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   10215
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   7
         Left            =   8880
         TabIndex        =   447
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   446
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   5
         Left            =   6600
         TabIndex        =   445
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   4
         Left            =   7800
         TabIndex        =   444
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   3
         Left            =   8880
         TabIndex        =   443
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   2
         Left            =   7800
         TabIndex        =   442
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   1
         Left            =   6600
         TabIndex        =   441
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   440
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox ChkObserva 
         Caption         =   "Mostrar observaciones del vencimiento"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   4920
         Width           =   3375
      End
      Begin VB.ComboBox cboCobro 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado.frx":D0B0
         Left            =   9000
         List            =   "frmListado.frx":D0BD
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4560
         Width           =   975
      End
      Begin VB.ComboBox cboCobro 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado.frx":D0D0
         Left            =   9000
         List            =   "frmListado.frx":D0DD
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4190
         Width           =   975
      End
      Begin VB.CheckBox chkApaisado 
         Caption         =   "Formato apaisado"
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   31
         Top             =   6960
         Width           =   1695
      End
      Begin VB.ComboBox cmbCuentas 
         Height          =   315
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4080
         Width           =   2775
      End
      Begin VB.CheckBox chkFormaPago 
         Caption         =   "Agrupar por forma pago"
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CheckBox chkNOremesar 
         Caption         =   "Solo marcados NO remesar"
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   1
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   0
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   20
         Left            =   3360
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   19
         Left            =   1080
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkEfectosContabilizados 
         Caption         =   "Mostrar riesgo"
         Height          =   255
         Left            =   8280
         TabIndex        =   19
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CheckBox ChkAgruparSituacion 
         Caption         =   "Agrupar por situacion vencimiento"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   188
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   6840
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   149
         Text            =   "Text1"
         Top             =   6480
         Width           =   2775
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   6480
         Width           =   735
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   145
         Text            =   "Text1"
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   144
         Text            =   "Text1"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4920
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Desglosar cliente"
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Top             =   6000
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Totalizar por fecha"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7920
         TabIndex        =   28
         Top             =   5985
         Width           =   1815
      End
      Begin VB.OptionButton optLCobros 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   7920
         TabIndex        =   26
         Top             =   5640
         Width           =   1335
      End
      Begin VB.OptionButton optLCobros 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   25
         Top             =   5640
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   7320
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   1
         Left            =   6360
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   7320
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   0
         Left            =   6360
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   36
         Text            =   "Text5"
         Top             =   3240
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   3600
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCobrosPendCli 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   32
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8760
         TabIndex        =   34
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   8400
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   5520
         TabIndex        =   417
         Top             =   6360
         Width           =   4335
         Begin VB.OptionButton optCuenta 
            Caption         =   "Nombre"
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   1
            Left            =   2400
            MaskColor       =   &H00404000&
            TabIndex        =   30
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton optCuenta 
            Caption         =   "Cuenta"
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   0
            Left            =   720
            MaskColor       =   &H00404000&
            TabIndex        =   29
            Top             =   180
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   5280
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Devuelto"
         Height          =   195
         Index           =   45
         Left            =   8280
         TabIndex        =   439
         Top             =   4600
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Recibido"
         Height          =   195
         Index           =   44
         Left            =   8280
         TabIndex        =   438
         Top             =   4230
         Width           =   705
      End
      Begin VB.Image ImageSe 
         Height          =   240
         Index           =   1
         Left            =   5880
         Picture         =   "frmListado.frx":D0F0
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image ImageSe 
         Height          =   240
         Index           =   0
         Left            =   5880
         Picture         =   "frmListado.frx":13942
         Top             =   1200
         Width           =   240
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   7800
         Top             =   5565
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   5520
         Top             =   5565
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Cuentas"
         Height          =   195
         Index           =   40
         Left            =   1200
         TabIndex        =   318
         Top             =   4155
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N�mero factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   37
         Left            =   6720
         TabIndex        =   281
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   15
         Left            =   5280
         TabIndex        =   280
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   14
         Left            =   5280
         TabIndex        =   279
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   36
         Left            =   6000
         TabIndex        =   278
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vencimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   35
         Left            =   240
         TabIndex        =   277
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   275
         Top             =   2205
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   20
         Left            =   3120
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   19
         Left            =   840
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   6840
         Width           =   240
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   0
         Left            =   840
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   151
         Top             =   6480
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   150
         Top             =   6840
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   148
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde dpto"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   147
         Top             =   4920
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta dpto"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   146
         Top             =   5280
         Width           =   945
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmListado.frx":1A194
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmListado.frx":209E6
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   143
         Top             =   4560
         Width           =   1245
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   1
         Left            =   6000
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   0
         Left            =   6000
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   54
         Top             =   5280
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   52
         Top             =   2685
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   5280
         TabIndex        =   50
         Top             =   2325
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   5280
         TabIndex        =   49
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmListado.frx":27238
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmListado.frx":2DA8A
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cobros pendientes clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   0
         Left            =   2520
         TabIndex        =   42
         Top             =   240
         Width           =   4890
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   3120
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   41
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   1245
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   3285
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   9600
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha c�lculo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   8400
         TabIndex        =   37
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   276
         Top             =   2205
         Width           =   615
      End
   End
   Begin VB.Frame FrameTransferencias 
      Height          =   3135
      Left            =   120
      TabIndex        =   235
      Top             =   0
      Width           =   4935
      Begin VB.CheckBox chkCartaAbonos 
         Caption         =   "Carta abonos"
         Height          =   255
         Left            =   480
         TabIndex        =   525
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   237
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   16
         Left            =   3360
         TabIndex        =   239
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   15
         Left            =   1200
         TabIndex        =   238
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   236
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   3480
         TabIndex        =   241
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   240
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   248
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   247
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   480
         TabIndex        =   246
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   2400
         TabIndex        =   245
         Top             =   1920
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   16
         Left            =   3120
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   15
         Left            =   960
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   2400
         TabIndex        =   244
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   480
         TabIndex        =   243
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   9
         Left            =   120
         TabIndex        =   242
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame FrameNorma57Importar 
      Height          =   6615
      Left            =   0
      TabIndex        =   529
      Top             =   0
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3480
         TabIndex        =   540
         Text            =   "Text1"
         Top             =   6120
         Width           =   3615
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   539
         Text            =   "Text1"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNoram57Fich 
         Height          =   375
         Left            =   9840
         Picture         =   "frmListado.frx":342DC
         Style           =   1  'Graphical
         TabIndex        =   537
         ToolTipText     =   "Leer"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdContabilizarNorma57 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7920
         TabIndex        =   536
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   42
         Left            =   9240
         TabIndex        =   534
         Top             =   6000
         Width           =   975
      End
      Begin MSComctlLib.ListView lwNorma57Importar 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   531
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serie"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "N� Fact"
            Object.Width           =   1677
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fec. fact."
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Orden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fec Cobro"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView lwNorma57Importar 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   533
         Top             =   3600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "N� Fact"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Motivo"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   5
         Left            =   1800
         Top             =   6120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   78
         Left            =   240
         TabIndex        =   541
         Top             =   6120
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Leer fichero bancario"
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
         Index           =   2
         Left            =   7680
         TabIndex        =   538
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos erroneos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   77
         Left            =   240
         TabIndex        =   535
         Top             =   3360
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos encontrados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   76
         Left            =   120
         TabIndex        =   532
         Top             =   720
         Width           =   2250
      End
      Begin VB.Label Label2 
         Caption         =   "Importar fichero norma 57"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   23
         Left            =   120
         TabIndex        =   530
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameCompensaciones 
      Height          =   5295
      Left            =   120
      TabIndex        =   320
      Top             =   0
      Width           =   6855
      Begin VB.CheckBox chkCompensa 
         Caption         =   "Dejar s�lo importe compensacion"
         Height          =   255
         Left            =   960
         TabIndex        =   329
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Frame FrameCambioFPCompensa 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   345
         Top             =   2400
         Width           =   6495
         Begin VB.TextBox txtDescFPago 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   3120
            TabIndex        =   346
            Text            =   "Text1"
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtFPago 
            Height          =   285
            Index           =   8
            Left            =   2160
            TabIndex        =   325
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Forma de pago vto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   49
            Left            =   120
            TabIndex        =   347
            Top             =   240
            Width           =   1590
         End
         Begin VB.Image imgFP 
            Height          =   240
            Index           =   8
            Left            =   1800
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cboCompensaVto 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   323
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox txtConcpto 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   328
         Text            =   "Text1"
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   339
         Text            =   "Text1"
         Top             =   4320
         Width           =   3615
      End
      Begin VB.CommandButton cmdContabCompensaciones 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4560
         TabIndex        =   330
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   22
         Left            =   5640
         TabIndex        =   331
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   337
         Text            =   "Text1"
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox txtConcpto 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   327
         Text            =   "Text1"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   335
         Text            =   "Text1"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   326
         Text            =   "Text1"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   324
         Text            =   "Text1"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   333
         Text            =   "Text1"
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   23
         Left            =   2280
         TabIndex        =   322
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   0
         Left            =   480
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Compensa sobre Vto."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   47
         Left            =   240
         TabIndex        =   342
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label Label6 
         Caption         =   "Pagos"
         Height          =   255
         Index           =   21
         Left            =   960
         TabIndex        =   341
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Cobros"
         Height          =   255
         Index           =   20
         Left            =   960
         TabIndex        =   340
         Top             =   3840
         Width           =   495
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "frmListado.frx":34CDE
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   46
         Left            =   240
         TabIndex        =   338
         Top             =   3600
         Width           =   885
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmListado.frx":3B530
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmListado.frx":41D82
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   45
         Left            =   240
         TabIndex        =   336
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   44
         Left            =   240
         TabIndex        =   334
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   2
         Left            =   1920
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   23
         Left            =   1920
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha contab."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   43
         Left            =   240
         TabIndex        =   332
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Contabilizaci�n compensaciones"
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
         Height          =   405
         Index           =   12
         Left            =   720
         TabIndex        =   321
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame FrameGastosTranasferencia 
      Height          =   3255
      Left            =   120
      TabIndex        =   467
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdGastosTransfer 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   471
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtVarios 
         Height          =   1005
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   469
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   470
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   35
         Left            =   3840
         TabIndex        =   472
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Transferencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   68
         Left            =   240
         TabIndex        =   475
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   67
         Left            =   1320
         TabIndex        =   474
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   66
         Left            =   240
         TabIndex        =   473
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Gastos transferencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   19
         Left            =   840
         TabIndex        =   468
         Top             =   360
         Width           =   3450
      End
   End
   Begin VB.Frame FrameOperAsegComunica 
      Height          =   5655
      Left            =   120
      TabIndex        =   510
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame FrameFraPendOpAseg 
         Height          =   1455
         Left            =   120
         TabIndex        =   522
         Top             =   2520
         Width           =   4815
         Begin VB.CheckBox chkVarios 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   524
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Solo asegurados"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   523
            Top             =   720
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Frame FrameSelEmpre1 
         Height          =   3015
         Left            =   120
         TabIndex        =   519
         Top             =   1920
         Width           =   4815
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   840
            TabIndex        =   520
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label2 
            Caption         =   "Empresas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   521
            Top             =   360
            Width           =   825
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   960
            Picture         =   "frmListado.frx":485D4
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   1320
            Picture         =   "frmListado.frx":4871E
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdOperAsegComunica 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   518
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   35
         Left            =   3600
         TabIndex        =   516
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   34
         Left            =   1200
         TabIndex        =   515
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   39
         Left            =   3840
         TabIndex        =   511
         Top             =   5040
         Width           =   975
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   1
         Left            =   120
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   35
         Left            =   3360
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   39
         Left            =   2880
         TabIndex        =   517
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   75
         Left            =   240
         TabIndex        =   514
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   34
         Left            =   960
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   360
         TabIndex        =   513
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   22
         Left            =   2310
         TabIndex        =   512
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.Frame FrameRecaudaEjec 
      Height          =   3975
      Left            =   120
      TabIndex        =   494
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdRecaudaEjecutiva 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2520
         TabIndex        =   499
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   19
         Left            =   840
         TabIndex        =   498
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2040
         TabIndex        =   507
         Text            =   "Text5"
         Top             =   2760
         Width           =   2715
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   18
         Left            =   840
         TabIndex        =   497
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2040
         TabIndex        =   504
         Text            =   "Text5"
         Top             =   2400
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   33
         Left            =   3120
         TabIndex        =   496
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   32
         Left            =   840
         TabIndex        =   495
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   38
         Left            =   3720
         TabIndex        =   500
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Recaudaci�n ejecutiva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   21
         Left            =   840
         TabIndex        =   509
         Top             =   360
         Width           =   3210
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   47
         Left            =   120
         TabIndex        =   508
         Top             =   2760
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   19
         Left            =   600
         Picture         =   "frmListado.frx":48868
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   46
         Left            =   120
         TabIndex        =   506
         Top             =   2445
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   74
         Left            =   120
         TabIndex        =   505
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   18
         Left            =   600
         Picture         =   "frmListado.frx":4F0BA
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   33
         Left            =   2880
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   37
         Left            =   2400
         TabIndex        =   503
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   502
         Top             =   1605
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   32
         Left            =   600
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vencimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   73
         Left            =   120
         TabIndex        =   501
         Top             =   1200
         Width           =   1590
      End
   End
   Begin VB.Frame FrameListaRecep 
      Height          =   4095
      Left            =   120
      TabIndex        =   365
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Justificante recepci�n"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   375
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   3
         Left            =   3600
         TabIndex        =   370
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   369
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cboListPagare 
         Height          =   315
         ItemData        =   "frmListado.frx":5590C
         Left            =   1920
         List            =   "frmListado.frx":55919
         Style           =   2  'Dropdown List
         TabIndex        =   373
         Top             =   3000
         Width           =   735
      End
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Desglosar vencimientos"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   374
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdListaRecpDocum 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   376
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Tal�n"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   372
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Pagare"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   371
         Top             =   2400
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   25
         Left            =   3600
         TabIndex        =   368
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   24
         Left            =   1560
         TabIndex        =   367
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   4080
         TabIndex        =   378
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ID recepci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   63
         Left            =   240
         TabIndex        =   437
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   33
         Left            =   2880
         TabIndex        =   436
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   32
         Left            =   600
         TabIndex        =   435
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "LLevados a banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   59
         Left            =   240
         TabIndex        =   416
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   58
         Left            =   240
         TabIndex        =   415
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   25
         Left            =   480
         TabIndex        =   380
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  recepci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   52
         Left            =   240
         TabIndex        =   379
         Top             =   600
         Width           =   1410
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   25
         Left            =   3360
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   24
         Left            =   2880
         TabIndex        =   377
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   24
         Left            =   1200
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado recepci�n documentos"
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
         Height          =   405
         Index           =   14
         Left            =   120
         TabIndex        =   366
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameDividVto 
      Height          =   2415
      Left            =   120
      TabIndex        =   403
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   406
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdDivVto 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   407
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   27
         Left            =   4200
         TabIndex        =   408
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "euros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   62
         Left            =   3240
         TabIndex        =   434
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   57
         Left            =   240
         TabIndex        =   409
         Top             =   960
         Width           =   5040
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Dividir vencimiento "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   16
         Left            =   240
         TabIndex        =   405
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   56
         Left            =   240
         TabIndex        =   404
         Top             =   720
         Width           =   5040
      End
   End
   Begin VB.Frame FrameReclama 
      Height          =   3615
      Left            =   120
      TabIndex        =   418
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   422
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2520
         TabIndex        =   432
         Text            =   "Text5"
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   421
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2520
         TabIndex        =   429
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   29
         Left            =   3960
         TabIndex        =   420
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   28
         Left            =   1440
         TabIndex        =   419
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdReclamas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   423
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   30
         Left            =   4320
         TabIndex        =   424
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   600
         TabIndex        =   433
         Top             =   2160
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   15
         Left            =   1200
         Picture         =   "frmListado.frx":55927
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   600
         TabIndex        =   431
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   61
         Left            =   240
         TabIndex        =   430
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   10
         Left            =   1200
         Picture         =   "frmListado.frx":5C179
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   428
         Top             =   1005
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   29
         Left            =   3720
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Historico reclamaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   17
         Left            =   480
         TabIndex        =   427
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   30
         Left            =   480
         TabIndex        =   426
         Top             =   1005
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   28
         Left            =   1200
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha reclamacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   60
         Left            =   240
         TabIndex        =   425
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameProgreso 
      Height          =   1935
      Left            =   3360
      TabIndex        =   45
      Top             =   2280
      Width           =   4095
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label lbl2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblPPAL 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameListadoCaja 
      Height          =   3495
      Left            =   120
      TabIndex        =   189
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkCaja 
         Caption         =   "Mostrar saldos arrastrados"
         Height          =   255
         Left            =   240
         TabIndex        =   317
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   206
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   205
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   211
         Text            =   "Text5"
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   210
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.CommandButton cmdCaja 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   207
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   3840
         TabIndex        =   208
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   12
         Left            =   3720
         TabIndex        =   200
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   191
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   213
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   212
         Top             =   2160
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   9
         Left            =   840
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   8
         Left            =   840
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   209
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   204
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   203
         Top             =   885
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   12
         Left            =   3480
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   11
         Left            =   840
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   202
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   6
         Left            =   240
         TabIndex        =   190
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame FrameCobroGenerico 
      Height          =   2295
      Left            =   120
      TabIndex        =   311
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   316
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   13
         Left            =   600
         TabIndex        =   314
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2040
         TabIndex        =   313
         Text            =   "Text5"
         Top             =   960
         Width           =   2955
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4080
         TabIndex        =   312
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta gen�rica para los vencimientos "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   42
         Left            =   240
         TabIndex        =   315
         Top             =   360
         Width           =   3330
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   13
         Left            =   240
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameFormaPago 
      Height          =   2415
      Left            =   120
      TabIndex        =   225
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdFormaPago 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   230
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   3960
         TabIndex        =   231
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   229
         Text            =   "Text1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   228
         Text            =   "Text1"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   227
         Text            =   "Text1"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   226
         Text            =   "Text1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado formas de pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   8
         Left            =   120
         TabIndex        =   234
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   240
         TabIndex        =   233
         Top             =   885
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   232
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   5
         Left            =   840
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   4
         Left            =   840
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameDevEfec 
      Height          =   2535
      Left            =   120
      TabIndex        =   214
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton optImpago 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   218
         Top             =   2040
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optImpago 
         Caption         =   "Fecha devolucion"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   217
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   14
         Left            =   3720
         TabIndex        =   216
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   215
         Top             =   990
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   3840
         TabIndex        =   220
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdEfecDev 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   219
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado efectos devueltos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   7
         Left            =   240
         TabIndex        =   224
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   223
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   14
         Left            =   3480
         Top             =   1012
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   13
         Left            =   840
         Top             =   1012
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   222
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   221
         Top             =   1005
         Width           =   615
      End
   End
   Begin VB.Frame FrameDpto 
      Height          =   3255
      Left            =   120
      TabIndex        =   163
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   172
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdDepto 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   171
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   167
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   166
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   165
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   164
         Text            =   "Text5"
         Top             =   1440
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado departamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   4
         Left            =   240
         TabIndex        =   186
         Top             =   480
         Width           =   4650
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   170
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   240
         TabIndex        =   169
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   168
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   7
         Left            =   840
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   6
         Left            =   840
         Top             =   1440
         Width           =   240
      End
   End
   Begin VB.Frame FrameAgentes 
      Height          =   2775
      Left            =   120
      TabIndex        =   153
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   162
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdAgente 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   161
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   155
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   157
         Text            =   "Text1"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   4
         Left            =   720
         Top             =   1560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado agentes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   5
         Left            =   240
         TabIndex        =   187
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   160
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   159
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   158
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   5
         Left            =   720
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame FramePrevision 
      Height          =   4935
      Left            =   120
      TabIndex        =   249
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   260
         Text            =   "Text1"
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optPrevision 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   259
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton optPrevision 
         Caption         =   "Fecha Vto"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   258
         Top             =   3240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkPrevision 
         Caption         =   "Gastos"
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   257
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkPrevision 
         Caption         =   "Pagos"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   256
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkPrevision 
         Caption         =   "Cobros"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   255
         Top             =   2760
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   18
         Left            =   3480
         TabIndex        =   254
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   253
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   264
         Text            =   "Text1"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   252
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   263
         Text            =   "Text1"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   251
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrevisionGastosCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   261
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4800
         TabIndex        =   262
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos imprevistos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   274
         Top             =   3840
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblPrevInd 
         Height          =   495
         Left            =   240
         TabIndex        =   273
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detallar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   272
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   32
         Left            =   240
         TabIndex        =   271
         Top             =   3240
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   270
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   35
         Left            =   2640
         TabIndex        =   269
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   268
         Top             =   2160
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   18
         Left            =   3240
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   17
         Left            =   1080
         Top             =   2160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   1
         Left            =   960
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   0
         Left            =   960
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   30
         Left            =   240
         TabIndex        =   267
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   266
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   240
         TabIndex        =   265
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado tesorer�a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   10
         Left            =   360
         TabIndex        =   250
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameRecepcionDocumentos 
      Height          =   4815
      Left            =   120
      TabIndex        =   348
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtCCost 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   356
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtDescCCoste 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   413
         Text            =   "Text1"
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3120
         TabIndex        =   411
         Text            =   "Text5"
         Top             =   3480
         Width           =   3195
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   355
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkAgruparCtaPuente 
         Caption         =   "Agrupa apuntes cta puente"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   354
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton cmdRecepDocu 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   357
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   23
         Left            =   5400
         TabIndex        =   358
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   363
         Text            =   "Text1"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtConcpto 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   353
         Text            =   "Text1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtConcpto 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   352
         Text            =   "Text1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   360
         Text            =   "Text1"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   351
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   350
         Text            =   "Text1"
         Top             =   960
         Width           =   3735
      End
      Begin VB.Image imgCCoste 
         Height          =   240
         Index           =   0
         Left            =   1680
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Centro de coste"
         Height          =   255
         Index           =   29
         Left            =   480
         TabIndex        =   414
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   28
         Left            =   480
         TabIndex        =   412
         Top             =   3480
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   14
         Left            =   1680
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   55
         Left            =   120
         TabIndex        =   410
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmListado.frx":629CB
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Haber"
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   364
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmListado.frx":6921D
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   51
         Left            =   120
         TabIndex        =   362
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "Debe"
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   361
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   50
         Left            =   120
         TabIndex        =   359
         Top             =   840
         Width           =   495
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   1
         Left            =   1560
         Picture         =   "frmListado.frx":6FA6F
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Height          =   405
         Index           =   13
         Left            =   480
         TabIndex        =   349
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame FrameAseg_Bas 
      Height          =   5655
      Left            =   120
      TabIndex        =   287
      Top             =   0
      Width           =   6375
      Begin VB.Frame FrameAsegAvisos 
         Caption         =   "Avisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   120
         TabIndex        =   466
         Top             =   4080
         Visible         =   0   'False
         Width           =   6015
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Siniestro"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   301
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Prorroga"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   300
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Falta de pago"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   299
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FrameForpa 
         Height          =   615
         Left            =   360
         TabIndex        =   400
         Top             =   4080
         Width           =   5775
         Begin VB.OptionButton optFP 
            Caption         =   "Descripci�n tipo pago"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   402
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton optFP 
            Caption         =   "Descripci�n forma pago"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   401
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame FrameASeg2 
         Height          =   855
         Left            =   1560
         TabIndex        =   397
         Top             =   3120
         Width           =   4575
         Begin VB.OptionButton optFecgaASig 
            Caption         =   "Fecha vencimiento"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   399
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optFecgaASig 
            Caption         =   "Fecha factura"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   398
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame FrOrdenAseg1 
         Height          =   855
         Left            =   120
         TabIndex        =   308
         Top             =   3120
         Width           =   5895
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   296
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   297
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "P�liza"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   298
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ordenar por"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   41
            Left            =   0
            TabIndex        =   310
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdAsegBascios 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   307
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   5040
         TabIndex        =   309
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   3120
         TabIndex        =   303
         Text            =   "Text5"
         Top             =   2280
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   3120
         TabIndex        =   302
         Text            =   "Text5"
         Top             =   2640
         Width           =   2715
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   12
         Left            =   1800
         TabIndex        =   295
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   11
         Left            =   1800
         TabIndex        =   294
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   22
         Left            =   4440
         TabIndex        =   292
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   21
         Left            =   1800
         TabIndex        =   288
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   11
         Left            =   1440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   12
         Left            =   1440
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   40
         Left            =   240
         TabIndex        =   306
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   840
         TabIndex        =   305
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   840
         TabIndex        =   304
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   22
         Left            =   4200
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   19
         Left            =   3600
         TabIndex        =   293
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha solicitud"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   39
         Left            =   240
         TabIndex        =   291
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   21
         Left            =   1440
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   18
         Left            =   840
         TabIndex        =   290
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ccc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   11
         Left            =   240
         TabIndex        =   289
         Top             =   480
         Width           =   5970
      End
   End
   Begin VB.Frame FrameGastosFijos 
      Height          =   3615
      Left            =   2640
      TabIndex        =   449
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   31
         Left            =   5040
         TabIndex        =   457
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdGastosFijos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   456
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtGastoFijo 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   452
         Text            =   "Text1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtDescGastoFijo 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   462
         Text            =   "Text1"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtGastoFijo 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   451
         Text            =   "Text1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtDescGastoFijo 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   459
         Text            =   "Text1"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox chkDesglosaGastoFijo 
         Caption         =   "Desglosar gastos"
         Height          =   255
         Left            =   240
         TabIndex        =   455
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   31
         Left            =   4800
         TabIndex        =   454
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   30
         Left            =   2160
         TabIndex        =   453
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   35
         Left            =   3840
         TabIndex        =   465
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   34
         Left            =   1200
         TabIndex        =   464
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image imgGastoFijo 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":762C1
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   17
         Left            =   720
         TabIndex        =   463
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image imgGastoFijo 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":7CB13
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gasto fijo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   65
         Left            =   120
         TabIndex        =   461
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   16
         Left            =   720
         TabIndex        =   460
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   31
         Left            =   4440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   30
         Left            =   1800
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha cargo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   64
         Left            =   120
         TabIndex        =   458
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado gastos fijos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   18
         Left            =   720
         TabIndex        =   450
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame frameListadoPagosBanco 
      Height          =   3855
      Left            =   120
      TabIndex        =   381
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkPagBanco 
         Caption         =   "Mostrar abonos"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   389
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkPagBanco 
         Caption         =   "Ordenado por fecha vencimiento"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   388
         Top             =   3000
         Width           =   3015
      End
      Begin VB.CommandButton cmdListadoPagosBanco 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   390
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   394
         Text            =   "Text1"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   385
         Text            =   "Text1"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   391
         Text            =   "Text1"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   384
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   27
         Left            =   3840
         TabIndex        =   387
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   26
         Left            =   1560
         TabIndex        =   386
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   4800
         TabIndex        =   392
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   27
         Left            =   600
         TabIndex        =   396
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   600
         TabIndex        =   395
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   4
         Left            =   1200
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   3
         Left            =   1200
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   54
         Left            =   240
         TabIndex        =   393
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha efecto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   53
         Left            =   360
         TabIndex        =   383
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   27
         Left            =   3480
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   26
         Left            =   1200
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado pagos por banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   15
         Left            =   240
         TabIndex        =   382
         Top             =   360
         Width           =   5370
      End
   End
   Begin VB.Frame FrameListRem 
      Height          =   4935
      Left            =   120
      TabIndex        =   173
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkRem 
         Caption         =   "Formato banco"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   448
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CheckBox chkTipoRemesa 
         Caption         =   "Talones"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   194
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkTipoRemesa 
         Caption         =   "Pagar�s"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   193
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkTipoRemesa 
         Caption         =   "Efectos"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   192
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame FrameOrdenRemesa 
         Height          =   975
         Left            =   360
         TabIndex        =   343
         Top             =   3000
         Width           =   4575
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Fecha vencimiento"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   196
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Descr. cuenta "
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   198
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Cuenta "
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   197
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Factura"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   195
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CheckBox chkRem 
         Caption         =   "Desglosar recibos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   199
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   184
         Top             =   1995
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   183
         Top             =   1995
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   182
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   181
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdListRem 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   201
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   185
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo remesa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   48
         Left            =   120
         TabIndex        =   344
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   24
         Left            =   2760
         TabIndex        =   180
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   179
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   2760
         TabIndex        =   178
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   720
         TabIndex        =   177
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A�o remesa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   176
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N�mero remesa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   23
         Left            =   120
         TabIndex        =   175
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado remesas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   3
         Left            =   240
         TabIndex        =   174
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame frpagosPendientes 
      Height          =   7215
      Left            =   120
      TabIndex        =   55
      Top             =   0
      Width           =   5415
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   360
         TabIndex        =   526
         Top             =   6120
         Width           =   4695
         Begin VB.OptionButton optMostraFP 
            Caption         =   "Forma de pago"
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   528
            Top             =   180
            Width           =   2055
         End
         Begin VB.OptionButton optMostraFP 
            Caption         =   "Tipo de pago"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   527
            Top             =   180
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.ComboBox cmbCuentas 
         Height          =   315
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   283
         Text            =   "Text1"
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   282
         Text            =   "Text1"
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CheckBox chkProv2 
         Caption         =   "Desglosar proveedor"
         Height          =   255
         Left            =   2400
         TabIndex        =   80
         Top             =   5400
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkProv 
         Caption         =   "Totalizar por fecha"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   79
         Top             =   5760
         Width           =   1935
      End
      Begin VB.OptionButton optProv 
         Caption         =   "Fecha vencimiento"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   77
         Top             =   5760
         Width           =   2175
      End
      Begin VB.OptionButton optProv 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   76
         Top             =   5400
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   1860
         TabIndex        =   63
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   65
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdPagosprov 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   64
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   59
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   58
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2400
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   2040
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   3720
         TabIndex        =   57
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   56
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cuentas"
         Height          =   195
         Index           =   41
         Left            =   240
         TabIndex        =   319
         Top             =   2955
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   38
         Left            =   240
         TabIndex        =   286
         Top             =   3480
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   240
         TabIndex        =   285
         Top             =   3885
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   284
         Top             =   4245
         Width           =   465
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   7
         Left            =   840
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   6
         Left            =   840
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   78
         Top             =   5160
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha c�lculo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   75
         Top             =   4680
         Width           =   1125
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   5
         Left            =   1560
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Pagos pendientes proveedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   4890
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   73
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   72
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   71
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   840
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   68
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   67
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   3420
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   66
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.Frame FrameCompensaAbonosCliente 
      Height          =   6735
      Left            =   120
      TabIndex        =   476
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton cmdVtoDestino 
         Height          =   375
         Index           =   1
         Left            =   240
         Picture         =   "frmListado.frx":83365
         Style           =   1  'Graphical
         TabIndex        =   492
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   9120
         TabIndex        =   490
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdVtoDestino 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmListado.frx":83D67
         Style           =   1  'Graphical
         TabIndex        =   488
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   8880
         TabIndex        =   487
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   484
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ListView lwCompenCli 
         Height          =   3975
         Left            =   240
         TabIndex        =   483
         Top             =   1560
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "N� Fact"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fec. fact."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Vto"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Forma pago"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cobro"
            Object.Width           =   2884
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abonos"
            Object.Width           =   2884
         EndProperty
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   481
         Text            =   "Text5"
         Top             =   1080
         Width           =   3675
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   480
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompensar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   8520
         TabIndex        =   478
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   36
         Left            =   9720
         TabIndex        =   477
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Imprimir hco compensacion"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   493
         Top             =   6240
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   72
         Left            =   8160
         TabIndex        =   491
         Top             =   5685
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "  Establecer vencimiento destino"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   489
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rectifca./Abono"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   71
         Left            =   8880
         TabIndex        =   486
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   70
         Left            =   7440
         TabIndex        =   485
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   69
         Left            =   240
         TabIndex        =   482
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1560
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Compensaci�n abonos cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   20
         Left            =   1680
         TabIndex        =   479
         Top             =   240
         Width           =   4890
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Public Opcion As Byte
    '1.- Cobros pendientes por cliente
    
    '3.- Reclamaciones por mail
    
    '4.- lISTADO agentes
    '5.- Departamentos
    
    '6.- Listado remesas
    
    '8.- Listado caja
    
    '9-  Devol remesas
    
    '10.- Listado formas de pago

    
    '11.- Transferencias PRovee   (o confirmings (domicilados o caixaconfirming)
    
    '12.- Listado previsional de gstos/pagos
    
    '13.- Transferencias ABONOS
    
    
    'Operaciones aseguradas
    '----------------------------
    '15.-  datos basicos
    '16.-  listado facturacion
    '17.-  Impagados asegurados
    
    
    '20.- Pregunta cuenta COBRO GENERICO
    '       La pongo aqui pq tengo implemntado todo
    
    
    '22.- Datos para la contabilizacion de las compensaciones
        
    '23.- Datos para la contbailiacion de la recpcion de documentos
    
    
    '24.-  Listado de documento(tal/pag) recibidos
    
    '25.-  Listado de pagos ordenados por banco  **** AHORA NO DEBERIA ENTRAR AQUI
    
    '26.-  Cancel remesa TAL/PAG.  Cando los importe no coinden. Solicitud cta y cc
    '27.-  Divide el vencimiento en dos vtos a partir del importe introducido en el text
        
        
    '30.-  Historico RECLAMACIONES
    '31.-   Gastos fijos
        
    '33.-  ASEGURADOS.  Listados avisos falta pago, avisos prorroga, aviso siniestro
        
    '34.-  Eliminar una recepcion de documentos, que ya ha sido contb con la puente
        
    '35.-  Gastos transferencias
        
    '36.-  Compensar ABONOS cobros
            
    '38.-  Recaudacion ejecutiva
        
    '39.-   Informe de comunicacion al seguro
    '40.-    Fras pendientes operaciones aseguradas
    
    '42.-   IMportar fichero norma 57 (recibos al cobro en ventanilla)
    
    '43.-   Confirmings
    '44.-   Caixaconfirming   igual que el de arriba
        
    '45.-   Listado cobros AGENTES -- >BACCHUS
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmD As frmDepartamentos
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmS As frmSerie
Attribute frmS.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim I As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub









Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function










Private Sub cboCobro_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub cboCompensaVto_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub


Private Sub Check3_Click()

End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkAgruparCtaPuente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub ChkAgruparSituacion_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub



Private Sub chkApaisado_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkCaja_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub







Private Sub chkCompensa_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkDesglosaGastoFijo_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkEfectosContabilizados_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkEmail_Click()
    If chkEmail.Value = 1 Then
        Label4(17).Caption = "Asunto"
    Else
        Label4(17).Caption = "Firmante"
    End If
End Sub

Private Sub chkEmail_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkFormaPago_KeyPress(KeyAscii As Integer)
KeyPressGral KeyAscii
End Sub


Private Sub chkLstTalPag_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkMarcarUtlRecla_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub


Private Sub chkNOremesar_KeyPress(KeyAscii As Integer)
KeyPressGral KeyAscii
End Sub



Private Sub ChkObserva_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkPagBanco_Click(Index As Integer)
    Me.chkPagBanco(1).Visible = chkPagBanco(0).Value = 1  'el de abono SOLO para "tipo herbelca"
    
End Sub

Private Sub chkPagBanco_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkPrevision_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkRem_Click(Index As Integer)
    Me.FrameOrdenRemesa.Visible = Me.chkRem(0).Value = 1
End Sub

Private Sub chkTipoRemesa_Click(Index As Integer)
    chkRem(1).Visible = chkTipoRemesa(0).Value = 0
End Sub



Private Sub chkTipPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub chkTipPagoRec_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub cmbCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub cmdAgente_Click()
    SQL = "SELECT * from agentes"
    Cad = ""
    RC = ""
    If txtAgente(5).Text <> "" Then
        Cad = " codigo >=" & txtAgente(5).Text
        RC = "Desde " & txtAgente(5).Text & " - " & txtDescAgente(5).Text
    End If
    If txtAgente(4).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & " codigo <=" & txtAgente(4).Text
        RC = RC & "      Hasta " & txtAgente(4).Text & " - " & txtDescAgente(4).Text
    End If
    
    If Cad <> "" Then Cad = " WHERE " & Cad
    
    SQL = SQL & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "DELETE from Usuarios.zpendientes where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    SQL = "INSERT INTO Usuarios.zpendientes (codusu,  numorden,  nomforpa) VALUES (" & vUsu.Codigo & ","
    CONT = 0
    While Not RS.EOF
        Cad = RS!Codigo & ",'" & DevNombreSQL(RS!Nombre) & "')"
        Cad = SQL & Cad
        Conn.Execute Cad
        CONT = CONT + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    If CONT = 0 Then
        MsgBox "Ning�n dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    Cad = "DesdeHasta= """ & Trim(RC) & """|"
    
    With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 9
            .Show vbModal
        End With
    
    
    
End Sub

Private Sub cmdAsegBascios_Click()
Dim B As Boolean
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    Select Case Opcion
    Case 15
        B = ListAseguBasico
    Case 16
        'Listado facturacion operaciones aseguradas
        B = ListAsegFacturacion
    
    Case 17
        'Impagados
        B = ListAsegImpagos
        
    Case 18
        B = ListAsegEfectos
    Case 33
        B = AvisosAseguradora
    End Select
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If B Then
        'Impimir.
        Select Case Opcion
        Case 15
            SQL = ""
            'Cuenta
            Cad = DesdeHasta("C", 11, 12)
            SQL = Trim(SQL & Cad)
            
            Cad = DesdeHasta("F", 21, 22, "Fec. solicitud:")
            If SQL <> "" Then Cad = SaltoLinea & Trim(Cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            SQL = SQL & Cad
            
            
            'Formulas
            Cad = "Cuenta= """ & SQL & """|"
            
            'Fecha imp
            Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 31 'Opcion informe
        Case 16
        
            SQL = ""
            'Cuenta
            Cad = DesdeHasta("C", 11, 12)
            SQL = Trim(SQL & Cad)
            If Me.optFecgaASig(0).Value Then
                Cad = DesdeHasta("F", 21, 22, "Fec. Fact:")
            Else
                Cad = DesdeHasta("F", 21, 22, "Fec. Vto:")
            End If
            If SQL <> "" Then Cad = SaltoLinea & Trim(Cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            SQL = SQL & Cad
            
            
            'Formulas
            Cad = "Cuenta= """ & SQL & """|"
            
            'Fecha imp
            Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 32 'Opcion informe
        
        Case 17
            SQL = ""
            'Cuenta
            Cad = DesdeHasta("C", 11, 12)
            SQL = Trim(SQL & Cad)
            
            Cad = DesdeHasta("F", 21, 22, "Fec. Vto:")
            If SQL <> "" Then Cad = SaltoLinea & Trim(Cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            SQL = SQL & Cad
            
            
            'Formulas
            Cad = "Cuenta= """ & SQL & """|"
            
            'Fecha imp
            Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 33 'Opcion informe
        
        Case 18
        
            SQL = ""
            'Cuenta
            Cad = DesdeHasta("C", 11, 12)
            SQL = Trim(SQL & Cad)
            
            Cad = DesdeHasta("F", 21, 22, "Fec. Vto:")
            If SQL <> "" Then Cad = SaltoLinea & Trim(Cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            SQL = SQL & Cad
            
            
            'Formulas
            Cad = "Cuenta= """ & SQL & """|"
            
            'Fecha imp
            Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 34 'Opcion informe
        Case 33
            SQL = ""
            'Cuenta
            Cad = DesdeHasta("C", 11, 12)
            SQL = Trim(SQL & Cad)
            

            Cad = Trim(DesdeHasta("F", 21, 22, "Fecha aviso: "))
            If SQL <> "" Then Cad = SaltoLinea & Cad
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            SQL = SQL & Cad
            
            
            'Formulas
            Cad = "Cuenta= """ & SQL & """|"
            
            'Fecha imp
            Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            
            
            If Me.optAsegAvisos(0).Value Then
                SQL = "falta de pago"
            ElseIf Me.optAsegAvisos(1).Value Then
                SQL = "prorroga"
            Else
                SQL = "siniestro"
            End If
            Cad = Cad & "Titulo= """ & SQL & """|"
            
            I = 3  'Numero parametros
            CONT = 90 'Opcion informe
        End Select
        
        
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = I
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = CInt(CONT)
            .Show vbModal
        End With
    
        
    End If
End Sub

Private Sub cmdCaja_Click()

    'Listado caja
    
        'Voy a comprobar , si tiene caja y ademas si la caja es
        ' el de predeterminado O NO
        I = vUsu.Codigo Mod 100
        SQL = "predeterminado"
        Cad = DevuelveDesdeBD("ctacaja", "susucaja", "codusu", CStr(I), "N", SQL)
        If Cad = "" And vUsu.Nivel > 0 Then
            MsgBox "Cajas sin asignar", vbExclamation
            Exit Sub
        End If
        
        If vUsu.Nivel > 0 Then
            If SQL = "1" Then
              'CAJA PRINCIPAL, las muestra todas
              SQL = ""
            Else
              SQL = " AND slicaja.codusu = " & vUsu.Codigo Mod 100
            End If
        Else
            SQL = ""
        End If
    
        RC = CampoABD(Text3(11), "F", "feccaja", True)
        If RC <> "" Then SQL = SQL & " AND " & RC
        RC = CampoABD(Text3(12), "F", "feccaja", False)
        If RC <> "" Then SQL = SQL & " AND " & RC
         
        RC = CampoABD(txtCta(9), "T", "ctacaja", True)
        If RC <> "" Then SQL = SQL & " AND " & RC
        
        RC = CampoABD(txtCta(10), "T", "ctacaja", False)
        If RC <> "" Then SQL = SQL & " AND " & RC
        
        SQL = " AND susucaja.codusu = slicaja.codusu " & SQL
               
        Set RS = New ADODB.Recordset
        
        Cad = "Select count(*) from slicaja,susucaja where slicaja.codusu>=0 " & SQL
                            'Pongo numlinea para asi no tener k comrpobar si es AND , where o su pu... madre
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        If Not RS.EOF Then
            If DBLet(RS.Fields(0), "N") > 0 Then I = 1
        End If
        RS.Close
        Set RS = Nothing
        If I = 0 Then
            MsgBox "Ningun registro con esos parametros", vbExclamation
            Exit Sub
        End If
        If chkCaja.Value = 1 Then
            I = 41  'saldo arrastrado
        Else
            I = 12  'normal
        End If
        
        If ImpirmirListadoCaja(SQL, Me.chkCaja.Value = 1) Then
            With frmImprimir
                .OtrosParametros = ""
                .NumeroParametros = 0
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = I
                .Show vbModal
            End With
        End If
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
End Sub

Private Sub cmdCanceRemTalPag_Click()
        
    'Esta visible la cta contable. Con lo cual es OBLIGADO ponerala
    If txtCta(14).Text = "" Then
        MsgBox "Debe indicar cta " & Label4(55).Caption, vbExclamation
        Exit Sub
    End If
    If vParam.autocoste Then
        If Me.txtCCost(0).Text = "" Then
            MsgBox "Indique el centro de coste", vbExclamation
            Exit Sub
        End If
    Else
        Me.txtCCost(0).Text = ""
    End If

    CadenaDesdeOtroForm = txtCta(14).Text & "|" & Me.txtCCost(0).Text & "|"
    Unload Me
End Sub

Private Sub cmdCobrosAgenLin_Click()

        
    Screen.MousePointer = vbHourglass
    If GenerarDatosListadoCobrosParcialesAgente Then
        Label3(50).Caption = ""
        Cad = "DH= """ & DevfrmCCtas & """|"
        CadenaDesdeOtroForm = "bacCobrosAgente.rpt"   'Lo cogera, si fuera o fuese necesario de la scryst
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{zpendientes.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 96
            .Show vbModal
        End With
        
        
        Cad = "Impresion correcta?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
            Screen.MousePointer = vbHourglass
            Label3(50).Caption = "Ajustando cobros "
            Label3(50).Refresh
            Conn.BeginTrans
            If RealizarProcesoUpdateCobrosAgente Then
                Conn.CommitTrans
                MsgBox "Proceso finalizado", vbInformation
            Else
                Conn.RollbackTrans
            End If
            
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCobrosPendCli_Click()
Dim Tot As Byte
Dim OpcionListado As Integer
    'Hago las comprobaciones
    If Text3(0).Text = "" Then
        MsgBox "Fecha c�lculo no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    
    If Me.ChkAgruparSituacion.Value = 1 And Me.chkFormaPago.Value = 1 Then
        MsgBox "No puede agrupar por forma pago y por situaci�n del vencimiento", vbExclamation
        Me.ChkAgruparSituacion.Value = 0
        Exit Sub
    End If
    
    If ChkAgruparSituacion.Value = 1 And Me.chkEfectosContabilizados.Value = 0 Then
        Cad = "Los efectos remesados ser�n mostrados igualmente"
        MsgBox Cad, vbExclamation
    
    End If
    
    
    
    'QUIEREN DETALLAR LAS CUENTAS
    CadenaDesdeOtroForm = ""
    If Me.cmbCuentas(0).ListIndex = 1 Then
        
        frmVarios.Opcion = 21
        CadenaDesdeOtroForm = Me.cmbCuentas(0).Tag
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            Me.cmbCuentas(0).ListIndex = 0
            Exit Sub
        Else
            
            Me.cmbCuentas(0).Tag = CadenaDesdeOtroForm
            GeneraComboCuentas
            Me.cmbCuentas(0).ListIndex = 2
        End If
    Else
        If Me.cmbCuentas(0).ListIndex = 2 Then CadenaDesdeOtroForm = Me.cmbCuentas(0).Tag
    End If
    
    Screen.MousePointer = vbHourglass
    If CobrosPendientesCliente(CadenaDesdeOtroForm) Then
        'Tesxto que iran
        SQL = "FECHA CALCULO: " & Text3(0).Text & "  "
        
        'Fecha fac
        Cad = DesdeHasta("F", 1, 2, "F.Factura:")
        SQL = SQL & Cad & " "
        
        'Fecha Vto
        Cad = DesdeHasta("F", 19, 20, "F.VTO:")
        SQL = SQL & Cad
        
        
        Cad = ""
        If Me.cboCobro(0).ListIndex > 0 Then
            Cad = Cad & "["
            If Me.cboCobro(0).ListIndex > 1 Then Cad = Cad & "SIN "
            Cad = Cad & "Recibido]"
        End If
        
        If Me.cboCobro(1).ListIndex > 0 Then
            Cad = Cad & "["
            If Me.cboCobro(1).ListIndex > 1 Then Cad = Cad & "SIN "
            Cad = Cad & "devuelto]"
        End If
        If Cad <> "" Then Cad = "   " & Cad
        SQL = SQL & Cad
        
        
        
        'Agente
        If txtAgente(0).Text <> "" Or txtAgente(1).Text <> "" Then
            Cad = "    AGENTE ("
            If txtAgente(0).Text <> "" And txtAgente(1).Text <> "" Then
                'Ha puesto los dos campos
                If txtAgente(0).Text <> txtAgente(1).Text Then
                    'SON DISTINTOS
                    Cad = Cad & txtAgente(0).Text & " hasta " & txtAgente(1).Text
                Else
                    Cad = Cad & txtAgente(0).Text & "  " & Me.txtDescAgente(0).Text
                    Cad = UCase(Cad)
                End If
            Else
                
                If txtAgente(0).Text <> "" Then Cad = Cad & " desde " & txtAgente(0).Text
                If txtAgente(1).Text <> "" Then Cad = Cad & " hasta " & txtAgente(1).Text
            End If
            Cad = Cad & ")"
            SQL = SQL & Cad
        End If
            
        RC = ""
        Cad = DesdeHasta("NF", 0, 1, "N� Factura:")
        RC = RC & Cad
        
        Cad = DesdeHasta("S", 0, 1, "Serie:")
        RC = RC & Cad
        
        
        
        
        
        
        
        
        
        If RC <> "" Then
            RC = SaltoLinea & Trim(RC)
            SQL = SQL & RC
        End If
        'Cuenta
        Cad = DesdeHasta("C", 1, 0)
        If Cad <> "" Then Cad = SaltoLinea & Trim(Cad)
        SQL = SQL & Cad
        
        
        'Si lleva la cuentas seleccionadas una a una, las pondremos en el encabezado
        If Me.cmbCuentas(0).ListIndex = 2 Then
            If Me.cmbCuentas(0).Tag <> "" Then
                RC = Me.cmbCuentas(0).Tag
                Cad = ""
                Do
                    I = InStr(1, RC, "|")
                    If I > 0 Then
                        If Cad <> "" Then Cad = Cad & ","
                        Cad = Cad & "  " & Mid(RC, 1, I - 1)
                        RC = Mid(RC, I + 1)
                    End If
                Loop Until I = 0
                If Cad <> "" Then
                    Cad = SaltoLinea & "Cuentas: " & Cad
                    SQL = SQL & Cad
                End If
            End If
        End If
        
       
        
        'Forma pago
        Cad = DesdeHasta("FP", 0, 1)
        If Cad <> "" Then Cad = SaltoLinea & Trim(Cad)
        SQL = SQL & Cad
        
        Cad = PonerTipoPagoCobro_(False, False)
        SQL = SQL & Cad
            
        'Si no solo NO remesar
        '---------------------
        If Me.chkNOremesar.Value = 1 Then SQL = Trim(SQL & "  SOLO marca no remesar.")
        
        'Formulas
        Cad = "Cuenta= """ & SQL & """|"
        
        'Fecha imp
        Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        
        
        RC = ""
        'Totaliza
        If Me.optLCobros(0).Value Then
            Tot = Abs(Check2.Value)
        Else
            Tot = Abs(Check1.Value)
        End If
        Cad = Cad & "Totalizar= " & Tot & "|"
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            
            'Para saber cual abro
            If Me.optLCobros(0).Value Then
                If Check2.Value = 1 Then
                    '.Opcion = 1
                    OpcionListado = 1
                Else
                    '.Opcion = 3  'Sin desglosar datos cliente
                    OpcionListado = 3
                End If
            Else
                '.Opcion = 2
                OpcionListado = 2
            End If
            
            
            
            'Si agrupa por tipo de situacion
            If Me.ChkAgruparSituacion.Value = 0 And Me.chkFormaPago.Value = 0 Then
                'Si ordena por cta o nombre
                If Me.optCuenta(1).Value Then OpcionListado = OpcionListado + 70
            Else
                If Me.ChkAgruparSituacion.Value = 1 Then
                    'por cuenta o nombre
                    If Me.optCuenta(1).Value Then
                        OpcionListado = OpcionListado + 73 'del 74 al  76
                    Else
                        OpcionListado = OpcionListado + 20
                    End If
                End If
                If Me.chkFormaPago.Value = 1 Then
                    'por cuenta o nombnre
                    If Me.optCuenta(1).Value Then
                        OpcionListado = OpcionListado + 76 'del 77 al  79
                    Else
                        OpcionListado = OpcionListado + 50
                    End If
                End If
            End If


            If Me.chkApaisado(0).Value = 1 Then OpcionListado = OpcionListado + 500
    
            .Opcion = OpcionListado
             .Show vbModal
        End With

    
    End If
    Me.FrameProgreso.Visible = False
    Screen.MousePointer = vbDefault
        
    
    
End Sub

Private Function PonerTipoPagoCobro_(ParaSelect As Boolean, EsReclamacion As Boolean) As String
Dim I As Integer
Dim Sele As Integer
Dim AUX As String
Dim Visibles As Byte

    AUX = ""
    Sele = 0
    Visibles = 0
    If Not EsReclamacion Then
        For I = 0 To Me.chkTipPago.Count - 1
            If Me.chkTipPago(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPago(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        AUX = AUX & ", " & I
                    Else
                        AUX = AUX & "-" & Me.chkTipPago(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then AUX = ""
        
    Else
        '************************
        'Reclamaciones
        
        For I = 0 To Me.chkTipPagoRec.Count - 1
            If Me.chkTipPagoRec(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPagoRec(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        AUX = AUX & ", " & I
                    Else
                        AUX = AUX & "-" & Me.chkTipPagoRec(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then AUX = ""
    End If
    If AUX <> "" Then
        AUX = Mid(AUX, 2)
        AUX = "(" & AUX & ")"
    End If
    PonerTipoPagoCobro_ = AUX
End Function



Private Sub cmdCompensar_Click()
    
    Cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
    If Cad = "" Then
        MsgBox "No esta configurada la aplicaci�n. Falta informe(10)", vbCritical
        Exit Sub
    End If
    Me.Tag = Cad
    
    Cad = ""
    RC = ""
    CONT = 0
    TotalRegistros = 0
    NumRegElim = 0
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
            If Trim(lwCompenCli.ListItems(I).SubItems(6)) = "" Then
                'Es un abono
                TotalRegistros = TotalRegistros + 1
            Else
                NumRegElim = NumRegElim + 1
            End If
        End If
        If Me.lwCompenCli.ListItems(I).Bold Then
            Cad = Cad & "A"
            If CONT = 0 Then CONT = I
            
            
            
        End If
    Next
    
    I = 0
    SQL = ""
    If Len(Cad) <> 1 Then
        'Ha seleccionado o cero o mas de uno
        If txtimpNoEdit(0).Text <> txtimpNoEdit(1).Text Then
            'importes distintos. Solo puede seleccionar UNO
            SQL = "Debe selecionar uno(y solo uno) como vencimiento destino"
            
        End If
    Else
        'Comprobaremos si el selecionado esta tb checked
        If Not lwCompenCli.ListItems(CONT).Checked Then
            SQL = "El vencimiento seleccionado no esta marcado"
        
        Else
            'Si el importe Cobro es mayor que abono, deberia estar
            Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
            If Importe <> 0 Then
                If Importe > 0 Then
                    'Es un abono
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) = "" Then SQL = "cobro"
                Else
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) <> "" Then SQL = "abono"
                End If
                If SQL <> "" Then SQL = "Debe marcar un " & SQL
            End If
            
        End If
    End If
    If TotalRegistros = 0 Or NumRegElim = 0 Then SQL = "Debe selecionar cobro(s) y abono(s)" & vbCrLf & SQL
        
    'Sep 2012
    'NO se pueden borrar las observaciones que ya estuvieran
    'RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
    If CONT > 0 Then
        'Hay seleccionado uno vto
        Set miRsAux = New ADODB.Recordset
        Cad = "text41csb,text42csb,text43csb,text51csb,text52csb,text53csb,text61csb,text62csb,text63csb,text71csb,text72csb,text73csb,text81csb,text82csb,text83csb"
        Cad = "Select " & Cad & " FROM scobro WHERE numserie ='" & lwCompenCli.ListItems(CONT).Text & "' AND codfaccl="
        Cad = Cad & lwCompenCli.ListItems(CONT).SubItems(1) & " AND fecfaccl='" & Format(lwCompenCli.ListItems(CONT).SubItems(2), FormatoFecha)
        Cad = Cad & "' AND numorden = " & lwCompenCli.ListItems(CONT).SubItems(3)
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If miRsAux.EOF Then
            SQL = SQL & vbCrLf & " NO se ha encontrado el veto. destino"
        Else
            'Vamos a ver cuantos registros son
            CadenaDesdeOtroForm = ""
            RC = "0"
            For I = 0 To 14
                If DBLet(miRsAux.Fields(I), "T") = "" Then
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux.Fields(I).Name & "|"
                    RC = Val(RC) + 1
                End If
            Next I
            
                
                
            'If TotalRegistros + NumRegElim > 15 Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
            If TotalRegistros + NumRegElim > Val(RC) Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    Else
        If MsgBox("Seguro que desea realizar la compensaci�n?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    Me.FrameCompensaAbonosCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameCompensaAbonosCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub





Private Sub cmdContabCompensaciones_Click()

    'COmprobaciones y leches
    If Me.txtConcpto(0).Text = "" Or txtDiario(0).Text = "" Or Text3(23).Text = "" Or _
        Me.txtConcpto(1).Text = "" Then
        MsgBox "Todos los campos de contabilizacion  son obligatorios", vbExclamation
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        If Me.txtCtaBanc(2).Text = "" Then
            MsgBox "Campo banco no puede estar vacio", vbExclamation
            Exit Sub
        End If
    Else
        If Me.txtFPago(8).Text <> "" Then
            RC = DevuelveDesdeBD("codforpa", "sforpa", "codforpa", txtFPago(8).Text, "N")
            If RC = "" Then
                MsgBox "No existe la forma de pago", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If FechaCorrecta2(CDate(Text3(23).Text), True) > 1 Then
        Ponerfoco Text3(23)
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        'No compensa sobre ningun vencimiento.
        'No puede marcar la opcion del importe
        If chkCompensa.Value = 1 Then
            MsgBox "'Dejar s�lo importe compensaci�n' disponible cuando compense sobre un vencimiento", vbExclamation
            Exit Sub
        End If
    End If

    'Cargamos la cadena y cerramos
    CadenaDesdeOtroForm = Me.txtConcpto(0).Text & "|" & Me.txtConcpto(1).Text & "|" & txtDiario(0).Text & "|" & Text3(23).Text & "|" & Me.txtCtaBanc(2).Text & "|" & DevNombreSQL(txtDescBanc(2).Text) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.txtFPago(8).Text & "|" & Me.cboCompensaVto.ItemData(Me.cboCompensaVto.ListIndex) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkCompensa.Value & "|"
    Unload Me
End Sub

Private Sub cmdContabilizarNorma57_Click()
    SQL = ""
    If Me.lwNorma57Importar(0).ListItems.Count = 0 Then SQL = SQL & "-Ningun vencimiento desde el fichero" & vbCrLf
    If Me.txtCtaBanc(5).Text = "" Then SQL = SQL & "-Cuenta bancaria" & vbCrLf
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    'La madre de las batallas
    'El sql que mando
    SQL = "(numserie ,codfaccl,fecfaccl,numorden ) IN (select ccost,pos,nomdocum,numdiari from tmpconext "
    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " and numasien=0 ) "
    'CUIDADO. El trozo 'from tmpconext  WHERE codusu' tiene que estar extamente ASI
    '  ya que en ver cobros, si encuentro esto, pong la fecha de vencimiento la del PAGO por
    ' ventanilla que devuelve el banco y contabilizamos en funcion de esa fecha
            
            
    Cad = Format(Now, "dd/mm/yyyy") & "|" & Me.txtCtaBanc(5).Text & " - " & Me.txtDescBanc(5).Text & "|0|"  'efectivo
    With frmVerCobrosPagos
        .ImporteGastosTarjeta_ = 0
        .OrdenacionEfectos = 3
        .vSQL = SQL
        .OrdenarEfecto = True
        .Regresar = False
        .ContabTransfer = False
        .Cobros = True
        .Tipo = 0
        .SegundoParametro = ""
        'Los textos
        .vTextos = Cad
        .CodmactaUnica = ""

        .Show vbModal
    End With

    
    'Borro haya cancelado o no
    LimpiarDelProceso
End Sub

Private Sub cmdDepto_Click()

    RC = ""
    Cad = ""
    If txtCta(7).Text <> "" Then
        Cad = " AND departamentos.codmacta >='" & txtCta(7).Text & "'"
        RC = "Desde " & txtCta(7).Text & " - " & DtxtCta(7).Text
    End If
    
    If txtCta(8).Text <> "" Then
        Cad = Cad & " AND departamentos.codmacta <='" & txtCta(8).Text & "'"
        RC = RC & "  hasta " & txtCta(8).Text & " - " & DtxtCta(8).Text
    End If


    
    SQL = "select departamentos.codmacta, nommacta,dpto,descripcion from departamentos,cuentas where cuentas.codmacta=departamentos.codmacta" & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = "DELETE from Usuarios.zpendientes where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    CONT = 0
    SQL = "INSERT INTO Usuarios.zpendientes (codusu,  numorden,codforpa,  nomforpa, codmacta, nombre) VALUES (" & vUsu.Codigo & ","
    While Not RS.EOF
        CONT = CONT + 1
        Cad = CONT & "," & RS!Dpto & ",'" & DevNombreSQL(RS!Descripcion) & "','" & RS!codmacta & "','" & DevNombreSQL(RS!Nommacta) & "')"
        Cad = SQL & Cad
        Conn.Execute Cad
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    If CONT = 0 Then
        MsgBox "Ning�n dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    Cad = "DesdeHasta= """ & Trim(RC) & """|"
    
    With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 10
            .Show vbModal
        End With
    
End Sub

Private Sub cmdDivVto_Click()
Dim Im As Currency

    'Dividira el vto en dos. En uno dejara el importe que solicita y en el otro el resto
    'Los gastos s quedarian en uno asi como el cobrado si diera lugar
    SQL = ""
    If txtImporte(1).Text = "" Then SQL = "Ponga el importe" & vbCrLf
    
    RC = RecuperaValor(CadenaDesdeOtroForm, 3)
    Importe = CCur(RC)
    Im = ImporteFormateado(txtImporte(1).Text)
    If Im = 0 Then
        SQL = "Importe no puede ser cero"
    Else
        If Importe > 0 Then
            'Vencimiento normal
            If Im > Importe Then SQL = "Importe superior al m�ximo permitido(" & Importe & ")"
            
        Else
            'ABONO
            If Im > 0 Then
                SQL = "Es un abono. Importes negativos"
            Else
                If Im < Importe Then SQL = "Importe superior al m�ximo permitido(" & Importe & ")"
            End If
        End If
        
    End If
    
    
    If SQL = "" Then
        Set RS = New ADODB.Recordset
        
        
        'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        I = -1
        RC = "Select max(numorden) from scobro WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Error. Vencimiento NO encontrado: " & CadenaDesdeOtroForm
        Else
            I = RS.Fields(0) + 1
        End If
        RS.Close
        Set RS = Nothing
        
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Ponerfoco txtImporte(1)
        Exit Sub
        
    Else
        SQL = "�Desea desdoblar el vencimiento con uno de : " & Im & " euros?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    'OK.  a desdoblar
    SQL = "INSERT INTO scobro (`numorden`,`gastos`,impvenci,`fecultco`,`impcobro`,`recedocu`,"
    SQL = SQL & "`tiporem`,`codrem`,`anyorem`,`siturem`,reftalonpag,"
    SQL = SQL & "`numserie`,`codfaccl`,`fecfaccl`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,`ctabanc2`,`text33csb`,`text41csb`,`text42csb`,`text43csb`,`text51csb`,`text52csb`,`text53csb`,`text61csb`,`text62csb`,`text63csb`,`text71csb`,`text72csb`,`text73csb`,`text81csb`,`text82csb`,`text83csb`,`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`obs`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,iban) "
    'Valores
    SQL = SQL & " SELECT " & I & ",NULL," & TransformaComasPuntos(CStr(Im)) & ",NULL,NULL,0,"
    SQL = SQL & "NULL,NULL,NULL,NULL,NULL,"
    SQL = SQL & "`numserie`,`codfaccl`,`fecfaccl`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,`ctabanc2`,`text33csb`,`text41csb`,`text42csb`,`text43csb`,`text51csb`,`text52csb`,`text53csb`,`text61csb`,`text62csb`,`text63csb`,`text71csb`,`text72csb`,`text73csb`,`text81csb`,`text82csb`,"
    'text83csb`,
    SQL = SQL & "'Div vto." & Format(Now, "dd/mm/yyyy hh:nn") & "'"
    SQL = SQL & ",`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`obs`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,iban FROM "
    SQL = SQL & " scobro WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
    Conn.BeginTrans
    
    'Hacemos
    CONT = 1
    If Ejecuta(SQL) Then
        'Hemos insertado. AHora updateamos el impvenci del que se queda
        If Im < 0 Then
            'Abonos
            SQL = "UPDATE scobro SET impvenci= impvenci + " & TransformaComasPuntos(CStr(Abs(Im)))
        Else
            'normal
            SQL = "UPDATE scobro SET impvenci= impvenci - " & TransformaComasPuntos(CStr(Im))
        End If
        
        SQL = SQL & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        SQL = SQL & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
        If Ejecuta(SQL) Then CONT = 0 'TODO BIEN ******
    End If
    'Si mal, volvemos
    If CONT = 1 Then
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        CadenaDesdeOtroForm = I
        Unload Me
    End If
    
    
End Sub

Private Sub cmdEfecDev_Click()
    'Listado de efectos devueltos
    SQL = ""
    RC = CampoABD(Text3(13), "F", "fechadev", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(14), "F", "fechadev", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    Set RS = New ADODB.Recordset
    
    RC = "SELECT count(*) from sefecdev where numorden>=0" & SQL
    RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then I = 1
    End If
    RS.Close
    Set RS = Nothing
    
    If I = 0 Then
        RC = "Ningun dato para mostrar"
        If SQL <> "" Then RC = RC & " con esos valores"
        MsgBox RC, vbExclamation
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    If ListadoEfectosDevueltos(SQL) Then
        
        Cad = DesdeHasta("F", 13, 14)
        If Cad <> "" Then Cad = "Fecha devoluci�n: " & Cad
        Cad = "Desde= """ & Trim(Cad) & """|"
        If Me.optImpago(0).Value Then
            I = 13
        Else
            I = 14
        End If
        
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFormaPago_Click()
    Cad = ""
    RC = CampoABD(txtFPago(4), "N", "codforpa", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtFPago(5), "N", "codforpa", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    I = 0
    If Cad <> "" Then
        I = 1
        'Forma pago
        SQL = ""
        RC = DesdeHasta("FP", 4, 5)
        SQL = "Cuenta= """ & Trim(RC) & """|"
    
    Else
        I = 0
        SQL = ""
    End If
    
        
    If ListadoFormaPago(Cad) Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = I
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 27
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdGastosFijos_Click()
    If ListadoGastosFijos() Then
        With frmImprimir
            SQL = "Detalla= " & Abs(Me.chkDesglosaGastoFijo.Value) & "|DH= """ & Cad & """|"
            
            
            .OtrosParametros = SQL
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 89
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdGastosTransfer_Click()

       If Me.txtImporte(2).Text = "" Then
            CadenaDesdeOtroForm = 0
       Else
            CadenaDesdeOtroForm = Me.txtImporte(2).Text
       End If
       Unload Me
End Sub

Private Sub cmdListadoPagosBanco_Click()
    If ListadoOrdenPago Then
    
        'Orden de pagos. Habran dos. El que devuelve la funcion de abajo y
        'el acabado en F que ir� ordenado por fecha dentro del grupo del banco
    
        CadenaDesdeOtroForm = DevuelveDesdeBD("informe", "scryst", "codigo", 7) 'Orden de pago a bancos
        
        If Me.chkPagBanco(0).Value = 1 Then
            If CadenaDesdeOtroForm = "" Then
                MsgBox "Falta registro 7 scryst", vbExclamation
                Exit Sub
            End If
            SQL = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 4)
            SQL = SQL & "F.rpt"
            RC = App.Path & "\InformesT\" & SQL
            If Dir(RC, vbArchive) = "" Then
                MsgBox "No existe el listado ordenado por fecha. Consulte soporte t�cnico" & vbCrLf & "El programa continuar�", vbExclamation
            Else
                CadenaDesdeOtroForm = SQL
            End If
        End If
        With frmImprimir
            .NumeroParametros = 1
            .FormulaSeleccion = "{zlistadopagos.codusu}=" & vUsu.Codigo
            
            .SoloImprimir = False
            .Opcion = 62
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdListaRecpDocum_Click()
Dim NomFile As String


    'Si marca la opcion de imprimir el justifacante de recepcion, el desglose tiene que estar marcado
    If chkLstTalPag(2).Value = 1 Then
        chkLstTalPag(1).Value = 1
        NomFile = DevuelveNombreInformeSCRYST(8, "Confir. recepcion tal�n")
        If NomFile = "" Then Exit Sub  'El msgbox ya lo da la funcion
        
    Else
        NomFile = ""
    End If
    
    If GeneraDatosTalPag Then
        
        RC = "FechaIMP= " & Format(Now, "dd/mm/yyyy") & "|Cuenta= "
    
        SQL = DesdeHasta("F", 24, 25, "F. Recep")
        If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
            'Solo uno seleccionado
            Cad = "Tal�n"
            If (chkLstTalPag(0).Value = 1) Then Cad = "Pagar�"
            SQL = Trim(SQL & Space(15) & "F. pago: " & Cad)
        End If
        
        
        Cad = DesdeHasta("NF", 2, 3, "Id. ")
        If Cad <> "" Then
            SQL = Trim(SQL & Space(15) & Cad)
        End If
        
        
        
        If cboListPagare.ListIndex >= 1 Then
            If cboListPagare.ListIndex = 1 Then
                Cad = "Llevadas a "
            Else
                Cad = "Pendientes de llevar"
            End If
            Cad = Cad & " banco"
            SQL = Trim(SQL & Space(15) & Cad)
        End If
        SQL = RC & """" & SQL & """|"
        

        
        CadenaDesdeOtroForm = NomFile   'Por si es el ersonalizable
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            If chkLstTalPag(3).Value = 1 Then
                'Si esta marcado la confirmacion recepcion
                If chkLstTalPag(2).Value = 1 Then
                    .Opcion = 87
                Else
                    .Opcion = 61
                End If
            Else
                .Opcion = 63
            End If
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdListRem_Click()
Dim B As Boolean
    '-------------------------------------
    'LISTADO REMESAS
    'Utilizaremos las tablas de informes
    ' ztesoreriacomun, ztmplibrodiario
    '------------------------------------

    'Comprobaciones iniciales
    If Me.chkTipoRemesa(0).Value = 0 And Me.chkTipoRemesa(1).Value = 0 And Me.chkTipoRemesa(2).Value = 0 Then
        MsgBox "Seleccione alg�n tipo de remesa", vbExclamation
        Exit Sub
    End If
    
    If Me.chkTipoRemesa(0).Value = 1 And Me.chkRem(1).Value = 1 Then
        MsgBox "Formato banco para talones / pagar�s", vbExclamation
        Exit Sub
    End If
    
    If chkRem(1).Value = 1 Then
        If Me.chkRem(0).Value = 1 Then MsgBox "Listado banco NO detalla vencimientos", vbExclamation
        chkRem(0).Value = 0
    End If
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If Me.chkRem(1).Value Then
        'FORMATO BANCO
        B = ListadoRemesasBanco
    Else
        'El de siempre
        B = ListadoRemesas
    End If
    If B Then
        With frmImprimir
            RC = "0"
            If Me.chkRem(1).Value = 1 Then
                RC = "1"
                I = 88
            Else
                I = 11
                If Me.chkRem(0).Value = 1 Then RC = "1"
            End If
            .OtrosParametros = "Mostrar= " & RC & "|"
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNoram57Fich_Click()

    If Me.lwNorma57Importar(0).ListItems.Count > 0 Or lwNorma57Importar(1).ListItems.Count > 0 Then
        SQL = "Ya hay un proceso . � Desea importar otro archivo?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    Me.cmdContabilizarNorma57.Visible = False
    
    cd1.FileName = ""
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    LimpiarDelProceso
    Me.Refresh
   
    If procesarficheronorma57 Then
        
        'El fichero que ha entrado es correcto.
        'Ahora vamos a buscar los vencimientos
        If BuscarVtosNorma57 Then
            
            'AHORA cargamos los listviews
            CargaLWNorma57 True   'los correctos 'Si es que hay
            
            'Los errores
            CargaLWNorma57 False
    
    
    
            Me.cmdContabilizarNorma57.Visible = Me.lwNorma57Importar(0).ListItems.Count > 0
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOperAsegComunica_Click()
Dim B As Boolean


    'Fecha hasta
    SQL = ""
    If Text3(35).Text = "" Then SQL = SQL & "-Fecha hasta obligatoria" & vbCrLf

    If Opcion = 39 Then
            
            RC = ""
            For I = 1 To Me.ListView3.ListItems.Count
                If Me.ListView3.ListItems(I).Checked Then RC = RC & "1"
            Next
            If RC = "" Then SQL = SQL & "-Seleccione alguna empresa" & vbCrLf
            
            If SQL <> "" Then
                SQL = "Campos obligatorios: " & vbCrLf & vbCrLf & SQL
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
    Else
    
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Comun para los dos
    SQL = "DELETE FROM Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    
    If Opcion = 39 Then
        B = ComunicaDatosSeguro_
        I = 92
        CONT = 0
        SQL = ""
    Else
        B = GeneraDatosFrasAsegurados
        If Me.chkVarios(1).Value = 1 Then
            'Resumido
            I = 94
        Else
            I = 93
        End If
    End If
    If B Then
            SQL = ""
            CONT = 0
            RC = ""
            If Opcion <> 39 Then If Me.chkVarios(0).Value = 1 Then SQL = "SOLO asegurados"
                

            If Me.Text3(34).Text <> "" Then RC = RC & "desde " & Text3(34).Text
            If Me.Text3(35).Text <> "" Then RC = RC & "     hasta " & Text3(35).Text
            If RC <> "" Then
                RC = Trim(RC)
                RC = "Fechas : " & RC
                SQL = Trim(SQL & "       " & RC)
            End If
            
            SQL = "pDH= """ & SQL & """|"
            CONT = CONT + 1
            
            If Me.Opcion = 40 Then
                '   True: De factura ALZIRA
                '   False: vto      HERBELCA
                
                '//En el rpt DeFactura : Alzira es 1 (fra)    y herbelca es 0 (vto)
                RC = Abs(vParam.FechaSeguroEsFra)
                SQL = SQL & "DeFactura= " & RC & "|"
                CONT = CONT + 1
            End If
    
            'Declaracion seguro
            'Cominicacion datos grupo
            If Opcion = 39 Then
                RC = ""
                For NumRegElim = 1 To Me.ListView3.ListItems.Count
                    If Me.ListView3.ListItems(NumRegElim).Checked Then
                        If RC <> "" Then RC = RC & SaltoLinea
                        RC = RC & Me.ListView3.ListItems(NumRegElim).Text
                    End If
                Next
        
                SQL = SQL & "Empresas= """ & RC & """|"
                CONT = CONT + 1
                
                
                RC = DevuelveDesdeBD("informe", "scryst", "codigo", 11) '
                If RC = "" Then
                    MsgBox "No esta configurada la aplicaci�n. Falta informe(11)", vbCritical
                    Exit Sub
                End If
                CadenaDesdeOtroForm = RC
                
            End If
        
        
    End If
    Screen.MousePointer = vbDefault
        
    
    If B Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = CONT
            .FormulaSeleccion = "{ztesoreriacomun.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    
    
        'rComunicaSeguro.rpt
    Else
        MsgBox "No se ha generado ning�n dato", vbExclamation
    End If
    
End Sub

Private Sub cmdPagosprov_Click()
    'Hago las comprobaciones
    If Text3(5).Text = "" Then
        MsgBox "Fecha calculo no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    
    
    
   'QUIEREN DETALLAR LAS CUENTAS
    CadenaDesdeOtroForm = ""
    If Me.cmbCuentas(1).ListIndex = 1 Then
        
        frmVarios.Opcion = 21
        CadenaDesdeOtroForm = Me.cmbCuentas(1).Tag
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            Me.cmbCuentas(1).ListIndex = 0
            Exit Sub
        Else
            
            Me.cmbCuentas(1).Tag = CadenaDesdeOtroForm
            GeneraComboCuentas
            Me.cmbCuentas(1).ListIndex = 2
        End If
    Else
        If Me.cmbCuentas(1).ListIndex = 2 Then CadenaDesdeOtroForm = Me.cmbCuentas(1).Tag
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    If PagosPendienteProv(CadenaDesdeOtroForm) Then
        'Tesxto que iran
        SQL = "FECHA CALCULO: " & Text3(5).Text & "  "
        
        'Fechas
        Cad = DesdeHasta("F", 3, 4)
        SQL = SQL & Cad
        
        'Cuenta
        Cad = DesdeHasta("C", 2, 3)
        If Cad <> "" Then Cad = SaltoLinea & Trim(Cad)
        SQL = SQL & Cad
        
        
        'Si lleva la cuentas seleccionadas una a una, las pondremos en el encabezado
        If Me.cmbCuentas(1).ListIndex = 2 Then
            If Me.cmbCuentas(1).Tag <> "" Then
                RC = Me.cmbCuentas(1).Tag
                Cad = ""
                Do
                    I = InStr(1, RC, "|")
                    If I > 0 Then
                        If Cad <> "" Then Cad = Cad & ","
                        Cad = Cad & "  " & Mid(RC, 1, I - 1)
                        RC = Mid(RC, I + 1)
                    End If
                Loop Until I = 0
                If Cad <> "" Then
                    Cad = SaltoLinea & "Cuentas: " & Cad
                    SQL = SQL & Cad
                End If
            End If
        End If
        
        
        
        
        
        
        'Desde hasta FP
        Cad = DesdeHasta("FP", 6, 7)
        If Cad <> "" Then Cad = SaltoLinea & Trim(Cad)
        SQL = SQL & Cad
        
        
        'Formulas
        Cad = "Cuenta= """ & SQL & """|"
        
        'Fecha imp
        Cad = Cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        
        
        
        'Totaliza
        Cad = Cad & "Totalizar= " & Abs(chkProv.Value) & "|"
        'marzo 2014
        Cad = Cad & "EsPorTipo= " & Abs(Me.optMostraFP(0).Value) & "|"
        
        
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 4
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            If Me.optProv(0).Value Then
                If chkProv2.Value Then
                    .Opcion = 4
                Else
                    .Opcion = 6
                End If
            Else
                .Opcion = 5
            End If

            
            .Show vbModal
        End With

    
    End If
    Me.FrameProgreso.Visible = False
    Screen.MousePointer = vbDefault
    
    
    

End Sub



Private Sub cmdPrevisionGastosCobros_Click()


    'Borramos las lineas en usuarios
    lblPrevInd.Caption = "Preparando ..."
    lblPrevInd.Refresh
    Conn.Execute "DELETE FROM Usuarios.ztmpconext WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM Usuarios.ztmpconextcab WHERE codusu =" & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset



    'Hacemos el selecet
    SQL = "select cuentas.codmacta,nommacta from ctabancaria,cuentas where cuentas.codmacta=ctabancaria.codmacta"
    RC = CampoABD(txtCtaBanc(0), "T", "ctabancaria.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCtaBanc(1), "T", "ctabancaria.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TotalRegistros = 0
    While Not RS.EOF
        '---
        If Not HacerPrevisionCuenta(RS!codmacta, RS!Nommacta) Then
        '---
            SQL = "DELETE FROM Usuarios.ztmpconextcab WHERE codusu =" & vUsu.Codigo
            SQL = SQL & " AND cta ='" & RS!codmacta & "'"
            Conn.Execute SQL
        Else
            TotalRegistros = TotalRegistros + 1
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    lblPrevInd.Caption = ""
    Me.Refresh
    
    
    If TotalRegistros = 0 Then
        MsgBox "Ningun dato generado", vbExclamation
        Exit Sub
    End If
    
    If Me.optPrevision(0).Value Then
        SQL = "Fecha"
    Else
        SQL = "Tipo"
    End If
    'txtCtaBanc  txtDescBanc
    
    
    
    SQL = "Titulo= ""Informe tesorer�a (" & SQL & ")""|"
    'Fechas intervalor
    SQL = SQL & "Fechas= ""Fecha hasta " & Text3(18).Text & """|"
    'Cuentas
    RC = DesdeHasta("BANCO", 0, 1)
    SQL = SQL & "Cuenta= """ & RC & """|"
    SQL = SQL & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
    SQL = SQL & "NumPag= 0|"
    SQL = SQL & "Salto= 2|"

    'SQL = SQL & "MostrarAnterior= " & MostrarAnterior & "|"
    
    Screen.MousePointer = vbDefault
    With frmImprimir
        .OtrosParametros = SQL
        .NumeroParametros = 6
        .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
        '.SoloImprimir = True
        'Opcion dependera del combo
        .Opcion = 29
        .Show vbModal
    End With
    

    
    
    
    
End Sub

Private Function HacerPrevisionCuenta(Cta As String, Nommacta As String) As Boolean
Dim SaldoArrastrado As Currency
Dim ID As Currency
Dim IH As Currency


    HacerPrevisionCuenta = False
    
    lblPrevInd.Caption = Cta & " - " & Nommacta
    lblPrevInd.Refresh
    ' Las fechas son del periodo, luego me importa una mierda las fechas desde hasta
    '
    '
    CargaDatosConExt Cta, Now, Now, " 1 = 1", Nommacta
    
    Conn.Execute "insert into Usuarios.ztmpconextcab select * from tmpconextcab where codusu =" & vUsu.Codigo
    
    Conn.Execute "DELETE FROM tmpfaclin where codusu =" & vUsu.Codigo
    
    RC = "INSERT INTO tmpfaclin (codusu, IVA,codigo, Fecha, Cliente, cta,"
    RC = RC & " ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
    
    'PARA CADA CUENTA
    'mETEREMOS TODOS LOS REGISTROS EN LA TABLA
    '
    '           TMPFACLIN
    '
    'TANTO COBROS COMO PAGOS I GASTOS
    '
    'Luego, en funcion del orden(TIPO o fecha) los iremos insertando en la tabla, para que
    'el saldo que va arrastrando sea el correcto
    
    
       
        
    CONT = 0
    
    
    '--------------------
    'DETALLAR COBROS
    lblPrevInd.Caption = Cta & " - Cobros"
    lblPrevInd.Refresh
    SQL = " WHERE fecvenci<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    SQL = SQL & " AND ctabanc1 ='" & Cta & "'"
    If chkPrevision(0).Value = 0 Then
        SQL = "select sum(impvenci),sum(impcobro),fecvenci from scobro " & SQL
        SQL = SQL & " GROUP BY fecvenci"
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        While Not miRsAux.EOF
        
            ID = DBLet(miRsAux.Fields(0), "N")
            IH = DBLet(miRsAux.Fields(1), "N")
            Importe = ID - IH

            If Importe <> 0 Then
                CONT = CONT + 1
                Cad = "'COBRO'," & CONT & ",'" & Format(miRsAux!fecvenci, FormatoFecha) & "','COBROS PENDIENTES',NULL,"
                'HAY COBROS
                If Importe < 0 Then
                    Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                Cad = RC & Cad & ")"
                Conn.Execute Cad
                
            End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
                
    Else
         'DETALLAR PAGOS COBROS
            '(codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
            'timporteD,timporteH, saldo
            
        'SQL = "select scobro.*,nommacta from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
        'SQL = SQL & " AND fecvenci<='2006-01-01'"
         
        SQL = "select scobro.* from scobro " & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            Cad = "'COBRO'," & CONT & ",'" & Format(miRsAux!fecvenci, FormatoFecha) & "','"
            'NUmero factura
            Cad = Cad & miRsAux!NUmSerie & miRsAux!codfaccl & "/" & miRsAux!numorden & "',"
            
            Cad = Cad & "'" & miRsAux!codmacta & "',"
            Importe = miRsAux!impvenci - DBLet(miRsAux!impcobro, "N")
            If Importe <> 0 Then
                If Importe < 0 Then
                    Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                Cad = Cad & ")"
                Cad = RC & Cad
                Conn.Execute Cad
            End If
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        
    End If
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR PAGOS
    '--------------------
    '--------------------
    lblPrevInd.Caption = Cta & " - pagos"
    lblPrevInd.Refresh
    SQL = " WHERE fecefect<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    SQL = SQL & " AND ctabanc1 ='" & Cta & "'"
    
    If chkPrevision(1).Value = 0 Then
        SQL = "select sum(impefect),sum(imppagad),fecefect from spagop " & SQL & " GROUP BY fecefect"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Importe = 0
        While Not miRsAux.EOF

                ID = DBLet(miRsAux.Fields(0), "N")
                IH = DBLet(miRsAux.Fields(1), "N")
                Importe = ID - IH
            
                If Importe <> 0 Then
                    CONT = CONT + 1
                    Cad = "'PAGO'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','PAGOS PENDIENTES',NULL,"
                    'HAY COBROS
                    If Importe > 0 Then
                        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                    Else
                        Cad = Cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                    End If
                    Cad = RC & Cad & ")"
                    Conn.Execute Cad
                End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
    Else
         'DETALLAR PAGOS COBROS
        'codusu, IVA,codigo, Fecha, Cliente, cta,"
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
        
        SQL = "select spagop.* from spagop " & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            Cad = "'PAGO'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','"
            'NUmero factura
            Cad = Cad & DevNombreSQL(miRsAux!numfactu) & "/" & miRsAux!numorden & "',"
            
            Cad = Cad & "'" & miRsAux!ctaprove & "',"
            Importe = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            If Importe <> 0 Then
                If Importe > 0 Then
                    Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                Else
                    Cad = Cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                End If
                Cad = Cad & ")"
                Cad = RC & Cad
                Conn.Execute Cad
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR GASTOS GASTOS
    '--------------------
    '--------------------
    
    SQL = " from sgastfij,sgastfijd where sgastfij.codigo= sgastfijd.codigo"
    SQL = SQL & " and fecha >='" & Format(Now, FormatoFecha)
    SQL = SQL & "' AND fecha <='" & Format(Format(Text3(18).Text, FormatoFecha), FormatoFecha) & "'"
    SQL = SQL & " and ctaprevista='" & Cta & "'"
    
    'Desde 5 Abril 2006
    '------------------
    ' Si el gasto esta contbilizado desde la tesoreria, tiene la marca "contabilizado"
    SQL = SQL & " and contabilizado=0"
    
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
     
     
    'ABro el recodset aqui.
    'Si es EOF entonces no necesito abrir la pantalla, puesto
    ' que no habran gastos para seleccionar
    'Si NO es EOF entonces abro el form y entonces alli(en frmvarios)
    'recorro el recodset
    SQL = " select sgastfij.codigo,descripcion,fecha,importe " & SQL
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If miRsAux.EOF Then
        miRsAux.Close
    Else
        NumRegElim = CONT
        CadenaDesdeOtroForm = "Gastos cuenta: " & Nommacta & "|" & Cta & "|" & Val(chkPrevision(2).Value) & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & RC & "|"
        frmVarios.Opcion = 18
        frmVarios.Show vbModal
        Set miRsAux = New ADODB.Recordset
        CONT = NumRegElim
        Me.Refresh
    End If
    
    
    If CONT = 0 Then Exit Function
    
    lblPrevInd.Caption = Cta & " - Informe"
    lblPrevInd.Refresh
    'Cargo INFORME
    '------------------------------------------------------------------------------------------
    'Leo el  saldo inicial
    RC = "Select * from tmpconextcab where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SaldoArrastrado = 0
    If Not miRsAux.EOF Then SaldoArrastrado = DBLet(miRsAux!acumtotT, "N")
    miRsAux.Close
    
    'Si desgloso cobros, los detallo, si no hago el acumu
    RC = "INSERT INTO Usuarios.ztmpconext (codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
    RC = RC & "timporteD,timporteH, saldo) VALUES (" & vUsu.Codigo & ",'" & Cta & "','"
        
    
    
    'Ahora cogere todos los registros que estan cargados en tmpfaclin y los metere ya
    'en la tabla con los importes, ordenado como dice el option y
    'arrastrando saldo
    SQL = "select tmpfaclin.*,nommacta from tmpfaclin left join cuentas on cta=codmacta where codusu =" & vUsu.Codigo & " ORDER BY "
    'EL ORDEN
    If optPrevision(0).Value Then
        SQL = SQL & "fecha,cta"
    Else
        SQL = SQL & "cta,fecha"
    End If
    CONT = 1
    ID = 0
    IH = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = Mid(miRsAux!iva, 1, 4) & "'," & CONT & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "','"
        
        
        
        If IsNull(miRsAux!Cta) Then
            'Stop
            Cad = Cad & "','" & DevNombreSQL(miRsAux!cliente) & "'"
        Else
            Cad = Cad & Mid(DevNombreSQL(miRsAux!cliente), 1, 10) & "',"
            If IsNull(miRsAux!Nommacta) Then
                Cad = Cad & "NULL"
            Else
                Cad = Cad & "'" & DevNombreSQL(miRsAux!Nommacta) & "'"
            End If
        End If
        If IsNull(miRsAux!Total) Then
            'VA AL DEBE
            Importe = miRsAux!impiva
            Cad = Cad & "," & TransformaComasPuntos(CStr(miRsAux!impiva)) & ",NULL,"
            ID = ID + Importe
        Else
            'HABER
            Importe = miRsAux!Total * -1
            Cad = Cad & ",NULL," & TransformaComasPuntos(CStr(miRsAux!Total)) & ","
            IH = IH + miRsAux!Total
        End If
        SaldoArrastrado = SaldoArrastrado + Importe
        Cad = Cad & TransformaComasPuntos(CStr(SaldoArrastrado)) & ")"
        Cad = RC & Cad
        Conn.Execute Cad
        
        
        CONT = CONT + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ajusto los importes de la tabla tmpconextcab
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumantD=acumtotD,acumantH=acumtotH,acumantT=acumtotT"
    SQL = SQL & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute SQL
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumperD=" & TransformaComasPuntos(CStr(ID))
    SQL = SQL & ", acumperH=" & TransformaComasPuntos(CStr(IH))
    SQL = SQL & ", acumperT=" & TransformaComasPuntos(CStr(ID - IH))
    SQL = SQL & ", acumtott=" & TransformaComasPuntos(CStr(SaldoArrastrado))
    
    SQL = SQL & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute SQL
    
    HacerPrevisionCuenta = True
    
End Function

Private Sub cmdRecaudaEjecutiva_Click()
    
    
    SQL = " scobro.codmacta=cuentas.codmacta AND"
    SQL = SQL & " fecejecutiva is null and impvenci+coalesce(gastos)-coalesce(impcobro,0)>0"
    'Si fechvto
    RC = CampoABD(Text3(32), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(33), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    'Codmacta
    RC = CampoABD(txtCta(18), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(18), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    
    'hacemos un COUNT
    RC = DevuelveDesdeBD("count(*)", "scobro,cuentas", SQL & " AND 1", "1")
    If RC = "" Then RC = "0"
    If Val(RC) = 0 Then
        MsgBox "No existen registros", vbExclamation
        Exit Sub
    End If
    
    SQL = " FROM scobro,cuentas WHERE " & SQL
    
    frmVarios.NumeroDocumento = SQL
    frmVarios.Opcion = 29
    frmVarios.Show vbModal
    
End Sub

Private Sub cmdRecepDocu_Click()
    If txtDiario(1).Text = "" Or Me.txtConcpto(2).Text = "" Or txtConcpto(3).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    If Me.Label4(55).Visible Then
        If Me.txtCta(14).Text = "" Then
            MsgBox "Cuentas " & Label4(55).Caption & " requerida", vbExclamation
            Exit Sub
        End If
        SQL = ""
        If vParam.autocoste Then
            RC = Mid(txtCta(14).Text, 1, 1)
            If RC = 6 Or RC = 7 Then
                If txtCCost(0).Text = "" Then
                    MsgBox "Centro de coste requerido", vbExclamation
                    Exit Sub
                Else
                    SQL = txtCCost(0).Text
                End If
            End If
            
                
        End If
        txtCCost(0).Text = SQL
        
    Else
        txtCCost(0).Text = ""
        Me.txtCta(14).Text = ""
    End If
    
    
    
    
    I = 0
    If Me.chkAgruparCtaPuente(0).Visible Then
        If Me.chkAgruparCtaPuente(0).Value Then I = 1
    End If
    CadenaDesdeOtroForm = txtDiario(1).Text & "|" & Me.txtConcpto(2).Text & "|" & txtConcpto(3).Text & "|" & I & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtCta(14).Text & "|" & txtCCost(0).Text & "|"
    
    Unload Me
End Sub

Private Sub cmdreclama_Click()
Dim NomArchivo As String
Dim Dpto As Integer
Dim EmpresaEscalona As Byte  '0 cualquiera  1 escalona
    
    
    SQL = ""
    
    'Si la fecha de reclamacion esta vacia----> mal
    If Text3(8).Text = "" Then SQL = SQL & "-Ponga la fecha de reclamaci�n" & vbCrLf
'        MsgBox "- Ponga la fecha de reclamaci�n", vbExclamation
'        Exit Sub
'    End If
    If txtDias.Text = "" Then SQL = SQL & "-Ponga los dias desde la ultima reclamaci�n" & vbCrLf
'        MsgBox "Ponga los dias desde la ultima reclamaci�n", vbExclamation
'        Exit Sub
'    End If
    
    If txtCarta.Text = "" Then SQL = SQL & "-Seleccione la carta a adjuntar" & vbCrLf
'        MsgBox "Seleccione la carta a adjuntar", vbExclamation
'        Exit Sub
'    End If
    
    
    'Si marca por email, NO puede marcar exlcuir clientes con email
    If chkEmail.Value = 1 Then
        If chkExcluirConEmail.Value = 1 Then SQL = SQL & "-En el envio de email no puede marcar la casilla 'excluir clientes con email'" & vbCrLf
'            MsgBox "En el envio de email no puede marcar la casilla 'excluir clientes con email'", vbExclamation
'            Exit Sub
'        End If
    End If
    
    If SQL <> "" Then
        SQL = "Opciones incorrectas: " & vbCrLf & vbCrLf & SQL
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    
    SQL = DevuelveDesdeBD("informe", "scryst", "codigo", 3) 'El tres es el tipo de docuemnto "reclamacion"

    If SQL = "" Then
            MsgBox "No existe la carta de reclamacion (3).", vbExclamation
            Exit Sub
    End If
    EmpresaEscalona = 0
    If LCase(Mid(SQL, 1, 3)) = "esc" Then EmpresaEscalona = 1
    
    NomArchivo = SQL
    SQL = App.Path & "\InformesT\" & SQL
    If Dir(SQL, vbArchive) = "" Then
        MsgBox "No se encuentra el archivo: " & SQL, vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Si poner marcar como reclamacion entonces debe estar marcada la opcion
    'de insertar en las tablas de col reclamas
    If chkMarcarUtlRecla.Value = 1 Then
        If Me.chkInsertarReclamas.Value = 0 Then
            MsgBox "Debe marcar tambien la opcion de ' INSERTAR REGISTROS RECLAMACIONES '", vbExclamation
            Exit Sub
        End If
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Ahora haremos todo el proceso
    I = Val(txtDias.Text)
    I = I * -1
    Fecha = CDate(Text3(8).Text)
    Fecha = DateAdd("d", I, Fecha)
    
    'Ya tenemos en F la fecha a partir de la cual reclamamos
    'Montamos el SQL
    MontaSQLReclamacion
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
    
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close
    
    If I = 0 Then
        MsgBox "No hay valores entre las fechas", vbExclamation
        Exit Sub
    End If
    
    'No enlazamos por NIF, si no k en NIF guardaremos codmacta
    
    

    'AHora empezamos con la generacion de datos
    'Borramos el anterior
    Cad = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad

    'Cadena insert
    Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos,"
    Cad = Cad & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, Asunto, contacto,Referencia) VALUES ("
    Cad = Cad & vUsu.Codigo
        
        
    'Monta Datos Empresa
    RS.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RS.EOF Then
        MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
        RC = ",'','','','','',''"  '6 campos
    Else
        RC = DBLet(RS!siglasvia) & " " & DBLet(RS!direccion) & "  " & DBLet(RS!numero) & ", " & DBLet(RS!puerta)
        RC = ",'" & DBLet(RS!nifempre) & "','" & vEmpresa.nomempre & "','" & RC & "','"
        RC = RC & DBLet(RS!codpos) & "','" & DBLet(RS!Poblacion) & "','" & DBLet(RS!provincia) & "'"
    End If
    RS.Close
    Cad = Cad & RC
    
    
    'Abrimos la carta
    RC = "SELECT * from scartas where codcarta = " & txtCarta.Text
    RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ' saludos, parrafo1, parrafo2, parrafo3,
    'parrafo4, parrafo5, despedida, Asunto, Referencia, contacto
    RC = ""
    For I = 2 To 6
        RC = RC & ",'" & DevNombreSQL(DBLet(RS.Fields(I))) & "'"
    Next I
    
    'Firmante , CArGO
    RC = RC & ",'" & txtVarios(0).Text & "','" & txtVarios(1).Text
    
    'Rc = Rc & "',NULL,NULL,NULL,NULL,NULL)"
    RC = RC & "',NULL,NULL,NULL)"
    Cad = Cad & RC
    'Cierro RS
    RS.Close
    
    
    'Insertamos carta
    Conn.Execute Cad
    
    'Para cada UNA la insertamos en la tmporal
    'Tomamos una tmp prestada
    'INSERT INTO zentrefechas (codusu, codigo, codccost, nomccost, conconam, nomconam,
    'codinmov, nominmov, fechaadq, valoradq, amortacu, fecventa, impventa, impperiodo) VALUES (
    Cad = "DELETE FROM USUARIOS.zentrefechas WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO USUARIOS.zentrefechas(codusu,codigo,codccost,nomccost,fecventa,conconam,fechaadq"
    SQL = SQL & ",nominmov,impventa,impperiodo,valoradq,codinmov) VALUES (" & vUsu.Codigo & ","
    
    'Nuevo. Febrero 2010. Departamento ira en codinmov
    
    'Codigo
    'Clave autonumerica
    '   codccost,nomccost,fecventa,conconam
    '    numserie,codfac,fecfac,numoreden
    '  Importes
    'en fechaadq pondremos codmacta, asi luego iremos a insertar
    
    I = 1
    While Not RS.EOF
    
        'Neuvo Febero 2010
        'Ademas de ver si me debe algo, si esta recibido NO lo puedo meter
        
        Importe = RS!impvenci + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
        If DBLet(RS!recedocu, "N") = 1 Then Importe = 0
        'If DBLet(Rs!recedocu, "N") = 1 And Importe > 0 Then Stop
        If Importe > 0 Then
            Cad = I & ",'" & RS!NUmSerie & "','"
            Cad = Cad & RS!codfaccl & "','"
            Cad = Cad & Format(RS!fecfaccl, FormatoFecha) & "',"
            Cad = Cad & RS!numorden & ",'"
            Cad = Cad & RS!codmacta & "','"
            'nomconam,impventa,impperiodo
            ' fec vto cobro, imp, cobrado
            Cad = Cad & RS!fecvenci & "',"
            Cad = Cad & TransformaComasPuntos(CStr(RS!impvenci)) & ","
            If IsNull(RS!impcobro) Then
                Cad = Cad & "NULL"
            Else
                Cad = Cad & TransformaComasPuntos(CStr(RS!impcobro))
            End If
            'ValorADQ=GASTOS
            Cad = Cad & "," & TransformaComasPuntos(CStr(DBLet(RS!Gastos, "N")))
            
            'Febrero 2010
            'Departamento
            Cad = Cad & "," & DBLet(RS!departamento, "N")
            Cad = SQL & Cad & ")"
            Conn.Execute Cad
            
            I = I + 1
            
        End If
        RS.MoveNext
        
    Wend
    RS.Close
    
    If I = 1 Then
        'Ningun valor con esa opcion
        MsgBox "No hay valores entre las fechas", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'AHora ya tenemos en entrefechas todos los valores k vamos a reclamar.
    'Para ello haremos un par cosas:
    ' 1.- Para cada codmacta(fechaadq) haremos su entrada en 347 cargando su datos NIF,dir,...
    ' 2.- UPDATEAREMOS nomconam con el NIF, para en el informe enalzar
    ' 3.- tabla cuentas. Donde guardaremos los datos de la cuenta bancaria
    
    Cad = "DELETE FROM Usuarios.z347  where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "DELETE FROM Usuarios.zcuentas  where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "SELECT fechaadq,codinmov FROM USUARIOS.zentrefechas WHERE codusu = " & vUsu.Codigo & " GROUP BY fechaadq,codinmov"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Datos contables
    Set miRsAux = New ADODB.Recordset
    CONT = 0
    While Not RS.EOF
        'BUSCAMOS DATOS
        Cad = "SELECT * from cuentas where codmacta='" & RS.Fields(0) & "'"
    
        'Insertar datos en z347
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Nuevo. Ya no llevamos NIF, llevaremos departamento
        RC = "" 'SERA EL NIF. Sera el DPTO
        I = 1
        If Not miRsAux.EOF Then
            'NIF -> codmacta
            RC = RS.Fields(0)
            Dpto = RS.Fields(1)
        Else
            'EOF
            I = 0
            MsgBox "No se encuentra la cuenta: " & RS.Fields(0), vbExclamation
            'NOS SALIMOS
            RS.Close
            Exit Sub
        End If
        
        'NO es EOF y tiene NIF
        If I > 0 Then
            'Aumentamos el contador
            CONT = CONT + 1
            
            
            'INSERTAMOS EN z347
            '-----------------------------------------
            SQL = "INSERT INTO Usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla,Provincia) "
            'Febrero 2010
            'SQL = SQL & "VALUES (" & vUsu.Codigo & ",0,'" & RC & "',0,'"
            SQL = SQL & "VALUES (" & vUsu.Codigo & "," & Dpto & ",'" & RC & "',0,'"
            
            
            'Razon social, dirdatos,codposta,despobla
            SQL = SQL & DevNombreSQL(DBLet(miRsAux!razosoci, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "','" & DBLet(miRsAux!codposta, "T") & "','" & DevNombreSQL(DBLet(miRsAux!despobla, "T"))
            SQL = SQL & "','" & DevNombreSQL(DBLet(miRsAux!desprovi, "T"))
            SQL = SQL & "')"
        
            Conn.Execute SQL
        
        
        
            
            SQL = "INSERT INTO Usuarios.zcuentas (codusu, codmacta, nommacta,despobla,razosoci,dpto) VALUES (" & vUsu.Codigo & ",'" & RC & "','"
            SQL = SQL & DBLet(miRsAux!nifdatos, "T") & "','" 'En nommacta meto el NIF del cliente
            If IsNull(miRsAux!Entidad) Then
                'Puede que sean todos nulos
                Cad = DBLet(miRsAux!oficina) & "   " & DBLet(miRsAux!CC, "T") & "    " & DBLet(miRsAux!cuentaba, "T")
                Cad = Trim(Cad)
            Else
                Cad = DBLet(miRsAux!IBAN, "T") & " " & Format(miRsAux!Entidad, "0000") & " " & Format(DBLet(miRsAux!oficina, "N"), "0000") & "  " & Format(DBLet(miRsAux!CC, "N"), "00") & " " & Format(DBLet(miRsAux!cuentaba, "N"), "0000000000")
            End If
            Cad = Cad & "','"
            'El dpto si tiene
            Cad = Cad & DevNombreSQL(DevuelveDesdeBD("descripcion", "departamentos", "codmacta = '" & miRsAux!codmacta & "' AND dpto", CStr(Dpto)))
            Cad = Cad & "'," & Dpto
            Ejecuta SQL & Cad & ")"   'Lo pongo en funcion para que no me de error
            
            
            'Updatear  FALTA### codusu = vusu.codusu
            SQL = "UPDATE USUARIOS.zentrefechas SET nomconam='" & RC & "' WHERE fechaadq = '" & RS!fechaadq & "'"
            SQL = SQL & " AND codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            
            
        End If
        miRsAux.Close
            
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
        
    If CONT = 0 Then
        MsgBox "Ningun dato devuelto para procesar por carta/mail", vbExclamation
        Exit Sub
    End If
    
    'Noviembre 2014
    'Comprobamos que todas las cuentas tienen email(si va por email)
    If Me.chkEmail.Value = 1 Then
            CadenaDesdeOtroForm = ""
            frmVarios.Opcion = 31
            frmVarios.Show vbModal
            
            If CadenaDesdeOtroForm = "" Then
                Screen.MousePointer = vbDefault
                Set RS = Nothing
                Exit Sub
            End If
    End If
    'AHORA YA ESTA. Si es carta, imprimimios directamente
    If chkEmail.Value = 0 Then
        'POR CARTA
        Cad = "FechaIMP= """ & Text3(8).Text & """|"
        Cad = Cad & "verCCC= " & Abs(Me.chkMostrarCta) & "|"
        CadenaDesdeOtroForm = NomArchivo
        
        
        With frmImprimir
            .EnvioEMail = False
            .OtrosParametros = Cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .Opcion = 7
            '
            .Show vbModal
        End With


    Else
        'POR MAIL. IREMOS UNO A UNO
        ' fechaadq = codmacta
        Screen.MousePointer = vbHourglass
        
        Cad = "DELETE FROM tmp347 WHERE codusu =" & vUsu.Codigo
        Conn.Execute Cad
        
        Cad = "SELECT fechaadq,maidatos,razosoci,nommacta FROM USUARIOS.zentrefechas,cuentas WHERE"
        Cad = Cad & " fechaadq=codmacta AND    CodUsu = " & vUsu.Codigo
        Cad = Cad & " GROUP BY fechaadq ORDER BY maidatos"
        RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        Cad = "FechaIMP= """ & Text3(8).Text & """|"
        Cad = Cad & "verCCC= " & Abs(Me.chkMostrarCta) & "|"
        SQL = "{ado.codusu}=" & vUsu.Codigo
        NumRegElim = 0
        CONT = 0
        frmPpal.Visible = False

        While Not RS.EOF
            Me.Refresh
            espera 0.5
            RC = DBLet(RS!maidatos, "T")
            If RC = "" Then
                
                If MsgBox("Sin mail para la cuenta: " & RS!fechaadq & " - " & RS!Nommacta & vbCrLf & "    �Continuar?", vbQuestion + vbYesNo) = vbNo Then
                    CONT = 0
                    RS.MoveLast
                End If
                
                SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe) VALUES (" & vUsu.Codigo
                SQL = SQL & ",0," & RS!fechaadq & ",NULL,0)"
                '
                'AL meter la cuenta con el importe a 0, entonces no la leera para enviarala
                'Pero despues si k podremos NO actualizar sus pagosya que no se han enviado nada
                Conn.Execute SQL
            Else
                Screen.MousePointer = vbHourglass
                With frmImprimir
                    CadenaDesdeOtroForm = NomArchivo
                    .OtrosParametros = Cad
                    .NumeroParametros = 1
                    SQL = "{ado.codusu}=" & vUsu.Codigo & " AND {ado.nif}= """ & RS.Fields(0) & """"
                    .FormulaSeleccion = SQL
                    .EnvioEMail = True
                    .QueEmpresaEs = EmpresaEscalona
                    .Opcion = 7
                    .Show vbModal
                    
                    If CadenaDesdeOtroForm = "OK" Then
                        Me.Refresh
                        espera 0.5
                        CONT = CONT + 1
                        'Se ha generado bien el documento
                        'Lo copiamos sobre app.path & \temp
                        SQL = RS.Fields(0) & ".pdf"
                        
                        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & SQL
                        
                        
                        'Insertamos en tmp347 la cuenta
                        SQL = "INSERT INTO tmp347(codusu, cliprov, cta, nif) VALUES (" & vUsu.Codigo & ",0,'" & RS.Fields(0) & "','" & SQL & "')"
                        Conn.Execute SQL
                        
                    End If
                    
                End With
            End If
            RS.MoveNext
        Wend
        RS.Close

        If CONT > 0 Then
             
             espera 0.5
             
             SQL = "Reclamacion fecha: " & Text3(8).Text & "|"
             
             SQL = SQL & "Reclamaci�n pago facturas efectuada el : " & Text3(8).Text & "|"
             
             'Escalona
             SQL = txtVarios(0).Text & "|Recuerde: En el archivo adjunto le enviamos informaci�n de su inter�s.|"

             frmEMail.queEmpresa = EmpresaEscalona
             frmEMail.Opcion = 3
             frmEMail.MisDatos = SQL
             frmEMail.Show vbModal
            
        End If
        
    End If
    
    Me.Hide
    frmPpal.Visible = True
    Me.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    'AHORA UPDATEAMOS LA FECHA RECLAMACION EN EL PAGO
    'SI ASI LO DESEA EL RECLAMANTE
    'Y SI SE HA REALIZADO; CUANTO MENOS; EL ENVIO
    '-----------------------------------------------------
    '-----------------------------------------------------
    If chkMarcarUtlRecla.Value = 1 Then
    
        'Si es por carta son todas, si es por mail, veremos si se ha llegado a enviar por mail, por lo menos
        'El mail sabemos k se ha enviado por que seran los k queden en tmp437
        'sin eliminar
        
        
        
        'Entonces veremos las reclamaciones k hemos efectuado bien, por email
        If Me.chkEmail.Value = 1 Then
            SQL = "SELECT * FROM tmp347 WHERE codusu=" & vUsu.Codigo & " AND Importe =0 "
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                'YA tengo la cuenta k no he podido enviar
                SQL = "DELETE from Usuarios.zentrefechas where codusu=" & vUsu.Codigo
                SQL = SQL & " AND nomconam = '" & RS!Cta & "'"
                Conn.Execute SQL
                'Siguiente
                RS.MoveNext
            Wend
            RS.Close
        End If
        
            
        'AHORA, las que queden en entrefechas seran las k he enviado por mail, con lo cual
        ' el proceso es el mismo k el de cartas
        
        SQL = "SELECT * from Usuarios.zentrefechas where codusu = " & vUsu.Codigo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = "UPDATE scobro set ultimareclamacion = '" & Format(Text3(8).Text, FormatoFecha) & "' WHERE numserie = '"
        While Not RS.EOF
            'VAMOS A MARCAR EL PAGO CON LA FECHA UTLMARECLAMCION
            Cad = RS!codccost & "' AND codfaccl = " & RS!nomccost & " AND fecfaccl  ='"
            Cad = Cad & Format(RS!fecventa, FormatoFecha) & "' AND numorden = " & RS!conconam
            Cad = SQL & Cad
            Conn.Execute Cad
    
            'Siguiente
            RS.MoveNext
        Wend
        RS.Close
    End If
    
    
    'FINALMENTE GRABAMOS LA TABLA HCO
    If chkInsertarReclamas.Value = 1 Then
        SQL = "SELECT MAX(codigo) from shcocob"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = 0
        If Not RS.EOF Then CONT = DBLet(RS.Fields(0), "N")
        RS.Close
        CONT = CONT + 1
    
        'INSERT INTO shcocob (codigo, numserie, codfaccl, fecfaccl, numorden, impvenci, codmacta, nommacta, carta) VALUES (
        SQL = "SELECT zentrefechas.*,nommacta from Usuarios.zentrefechas,cuentas where codusu = " & vUsu.Codigo
        SQL = SQL & " AND zentrefechas.nomconam=cuentas.codmacta"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = "INSERT INTO shcocob (fecreclama,carta,codigo, numserie, codfaccl, fecfaccl, numorden, impvenci, codmacta, nommacta)"
        SQL = SQL & " VALUES ('" & Format(Text3(8).Text, FormatoFecha) & "',"
        If Me.chkEmail.Value = 1 Then
            SQL = SQL & "1,"
        Else
            SQL = SQL & "0,"
        End If
        While Not RS.EOF
            Cad = CONT & ",'" & RS!codccost & "'," & RS!nomccost & ",'" & Format(RS!fecventa, FormatoFecha) & "',"
            Importe = RS!impventa + RS!valoradq - DBLet(RS!impperiodo, "N")
            Cad = Cad & RS!conconam & "," & TransformaComasPuntos(CStr(Importe)) & ",'"
            Cad = Cad & RS!nomconam & "','" & DevNombreSQL(RS!Nommacta) & "')"
            Cad = SQL & Cad
            Conn.Execute Cad
            'siguiente
            CONT = CONT + 1
            RS.MoveNext
        Wend
        RS.Close
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub MontaSQLReclamacion()
    
    'Siempre hay que a�adir el AND
    
    
    SQL = ""
    
    
    'Fecha factura
    RC = CampoABD(txtSerie(2), "T", "scobro.numserie", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtSerie(3), "T", "scobro.numserie", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    'Fecha factura
    RC = CampoABD(Text3(6), "F", "fecfaccl", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    RC = CampoABD(Text3(7), "F", "fecfaccl", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Fecha vto
    RC = CampoABD(Text3(9), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(Text3(10), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'cuenta
    RC = CampoABD(txtCta(4), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(5), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC

    
    
    'Agente
    RC = CampoABD(txtAgente(3), "N", "scobro.agente", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtAgente(2), "N", "scobro.agente", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    'Forma de pago
    RC = CampoABD(txtFPago(3), "N", "scobro.codforpa", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtFPago(2), "N", "scobro.codforpa", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Solo devueltos
    If chkReclamaDevueltos.Value = 1 Then SQL = SQL & " AND devuelto = 1"
      
    
    'Marzo2015
    If chkExcluirConEmail.Value = 1 Then SQL = SQL & " AND coalesce(maidatos,'')=''"
    
    
    'LA de la fecha
    SQL = SQL & " AND ((ultimareclamacion  is null) OR (ultimareclamacion <= '" & Format(Fecha, FormatoFecha) & "'))"
    
    'QUE FALTE POR PAGAR
    SQL = SQL & " AND (impvenci>0)"
    
    
    RC = PonerTipoPagoCobro_(True, True)
    If RC <> "" Then SQL = SQL & " AND tipforpa IN " & RC
    
    
    
    'Select
    Cad = "Select scobro.*, cuentas.codmacta FROM scobro,cuentas,sforpa "
    Cad = Cad & " WHERE  sforpa.codforpa=scobro.codforpa AND scobro.codmacta = cuentas.codmacta"
    Cad = Cad & " AND sforpa.codforpa=scobro.codforpa "
    SQL = Cad & SQL
    
    
    
    
    
End Sub






Private Sub cmdReclamas_Click()
    
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If ListadoReclamas Then
        With frmImprimir
            Cad = "Cadena= """ & Cad & """|"
            .OtrosParametros = Cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 86
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdTransfer_Click()
    
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If ListadoTransferencias Then
        With frmImprimir
            Cad = "Mostrar= 1|tipot= """
            I = 28
            SQL = "Listado transferencias"
            If Opcion = 11 Then
                Cad = Cad & "(Pagos)"
            ElseIf Opcion = 13 Then
                Cad = Cad & "(Abonos)"
                If Me.chkCartaAbonos.Value Then
                    CadenaDesdeOtroForm = DevuelveNombreInformeSCRYST(12, "Carta abono0")
                    I = 95
                End If
            Else
                If Opcion = 44 Then
                    SQL = "Caixa confirming"
                Else
                    SQL = "Pagos domiciliados"
                End If
            End If
            
            Cad = Cad & """|ErTitulo= """ & SQL & """|"
            
            .OtrosParametros = Cad
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            
            .Opcion = I
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVtoDestino_Click(Index As Integer)
    
    If Index = 0 Then
        TotalRegistros = 0
        If Not Me.lwCompenCli.SelectedItem Is Nothing Then TotalRegistros = Me.lwCompenCli.SelectedItem.Index
    
    
        For I = 1 To Me.lwCompenCli.ListItems.Count
            If Me.lwCompenCli.ListItems(I).Bold Then
                Me.lwCompenCli.ListItems(I).Bold = False
                Me.lwCompenCli.ListItems(I).ForeColor = vbBlack
                For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                    Me.lwCompenCli.ListItems(I).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwCompenCli.ListItems(I).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalRegistros > 0 Then
            I = TotalRegistros
            Me.lwCompenCli.ListItems(I).Bold = True
            Me.lwCompenCli.ListItems(I).ForeColor = vbRed
            For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                Me.lwCompenCli.ListItems(I).ListSubItems(CONT).ForeColor = vbRed
                Me.lwCompenCli.ListItems(I).ListSubItems(CONT).Bold = True
            Next
        End If
        lwCompenCli.Refresh
        
        PonerfocoObj Me.lwCompenCli

    Else
    
        Cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
        If Cad = "" Then
            MsgBox "No esta configurada la aplicaci�n. Falta informe(10)", vbCritical
            Exit Sub
        End If
        CadenaDesdeOtroForm = Cad
    
        LanzaBuscaGrid 1, 4


    End If
End Sub

Private Sub Command1_Click()
    If txtCta(13).Text = "" Then
        MsgBox "Ponga la cuenta", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = txtCta(13).Text
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 1
            Text3(1).SetFocus
        Case 3
            
            'Reclamaciones. Si no tiene configurado el envio web
            'no habilitaremos el check
            Cad = DevuelveDesdeBD("smtpHost", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
            If Cad = "" Then
                Me.chkEmail.Value = 0
                chkEmail.Enabled = False
            End If
            'Text3(6).SetFocus
            txtSerie(2).SetFocus
        Case 10
            Me.cmdFormaPago.SetFocus
        Case 12
            txtCtaBanc(0).SetFocus
        Case 20
            Ponerfoco txtCta(13)
            
        Case 22
            'Contabi efectos
            If CONT > 0 Then
                For I = 1 To Me.cboCompensaVto.ListCount
                    If Me.cboCompensaVto.ItemData(I) = CONT Then
                        CONT = I
                        Exit For
                    End If
                Next
            End If
            Me.cboCompensaVto.ListIndex = CONT
            Ponerfoco Text3(23)
        Case 23
            CadenaDesdeOtroForm = ""  'Para que  no devuelva nada
        Case 30
            Ponerfoco Text3(28)
            
        Case 31
            'gastos fijos
            Text3(30).Text = "01/01/" & Year(Now)
        Case 35
            Ponerfoco txtImporte(2)
            
        Case 36
            If CadenaDesdeOtroForm <> "" Then
                txtCta(17).Text = CadenaDesdeOtroForm
                txtCta_LostFocus 17
            Else
                Ponerfoco txtCta(17)
            End If
            CadenaDesdeOtroForm = ""
            
        Case 39
            Ponerfoco Text3(34)
            
        Case 42
            
'            Me.Refresh
'            cmdNoram57Fich_Click


        Case 45
            HabilitarCobrosParciales
            Ponerfoco Text3(46)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Me.imgCtaBanc, 1, "Cuenta contable bancaria"
    CargaImagenesAyudas Image2, 2
    CargaImagenesAyudas Me.imgFP, 1, "Forma de pago"
    CargaImagenesAyudas Me.Image3, 1, "Cuenta contable"
    CargaImagenesAyudas Me.Imagente, 1, "Seleccionar agente"
    CargaImagenesAyudas imgCCoste, 1, "Centro de coste"
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    For Each Img In Me.imgConcepto
        Img.ToolTipText = "Concepto"
    Next
    For Each Img In Me.imgDiario
        Img.ToolTipText = "Diario"
    Next
    
    
    
    For Each Img In Me.imgDpto
        Img.ToolTipText = "Departamento"
    Next
    
    
    'Limpiamos el tag
    txtCta(6).Tag = ""
    PrimeraVez = True
    FrCobrosPendientesCli.Visible = False
    frpagosPendientes.Visible = False
    FramereclaMail.Visible = False
    FrameAgentes.Visible = False
    FrameDpto.Visible = False
    FrameListRem.Visible = False
    FrameListadoCaja.Visible = False
    FrameDevEfec.Visible = False
    Me.FrameFormaPago.Visible = False
    FrameTransferencias.Visible = False
    Me.FramePrevision.Visible = False
    FrameAseg_Bas.Visible = False
    FrameCobroGenerico.Visible = False
    FrameCompensaciones.Visible = False
    FrameRecepcionDocumentos.Visible = False
    FrameListaRecep.Visible = False
    frameListadoPagosBanco.Visible = False
    FrameDividVto.Visible = False
    FrameReclama.Visible = False
    FrameGastosFijos.Visible = False
    FrameGastosTranasferencia.Visible = False
    FrameCompensaAbonosCliente.Visible = False
    FrameRecaudaEjec.Visible = False
    FrameOperAsegComunica.Visible = False
    FrameNorma57Importar.Visible = False
    FrameCobrosAgentesLin.Visible = False
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
    Select Case Opcion
    Case 1
        'Leeo la opcion del fichero x defecto
        Me.optCuenta(0).Value = CheckValueLeer("Listcta") = 1
        If Me.optCuenta(0).Value = False Then Me.optCuenta(1).Value = True
    
        chkApaisado(0).Value = Abs(CheckValueLeer("Infapa") = 1)
    
        Me.Frame1.BorderStyle = 0 'sin borde
        FrCobrosPendientesCli.Visible = True
        W = Me.FrCobrosPendientesCli.Width
        H = Me.FrCobrosPendientesCli.Height + 120
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        'Fecha = CDate(DiasMes(Month(Now), Year(Now)) & "/" & Month(Now) & "/" & Year(Now))
        'Text3(2).Text = Format(Fecha, "dd/mm/yyyy")
        Me.cmbCuentas(0).Tag = ""
        GeneraComboCuentas
        Me.cmbCuentas(0).ListIndex = 0
        
        Me.cboCobro(0).ListIndex = 2
        Me.cboCobro(1).ListIndex = 0
    Case 2
        frpagosPendientes.Visible = True
        W = Me.frpagosPendientes.Width
        H = Me.frpagosPendientes.Height
        Text3(5).Text = Format(Now, "dd/mm/yyyy")
        'Fecha = CDate(DiasMes(Month(Now), Year(Now)) & "/" & Month(Now) & "/" & Year(Now))
        'Text3(4).Text = Format(Fecha, "dd/mm/yyyy")
        Me.cmbCuentas(1).Tag = ""
        GeneraComboCuentas
        Me.cmbCuentas(1).ListIndex = 0
    Case 3
        Caption = "Reclamaciones"
        FramereclaMail.Visible = True
        W = Me.FramereclaMail.Width
        H = Me.FramereclaMail.Height
        Text3(8).Text = Format(Now, "dd/mm/yyyy")

        'ESPECIAL
        'Si no existe la carpeta tmp en app.path la creo
        If Dir(App.Path & "\temp", vbDirectory) = "" Then MkDir App.Path & "\temp"
        CargaTextosTipoPagos True
    Case 4
        
        Caption = "Agentes"
        FrameAgentes.Visible = True
        W = Me.FrameAgentes.Width
        H = Me.FrameAgentes.Height
        
    Case 5
         
        Caption = "Departamentos"
        FrameDpto.Visible = True
        W = Me.FrameDpto.Width
        H = Me.FrameDpto.Height
        
        
    Case 6, 7
         
        Caption = "Remesas"
        FrameListRem.Visible = True
        W = Me.FrameListRem.Width
        H = Me.FrameListRem.Height
        FrameOrdenRemesa.Visible = False
        
    Case 8
        FrameListadoCaja.Visible = True
        Caption = "Listado"
        W = Me.FrameListadoCaja.Width
        H = Me.FrameListadoCaja.Height
        
        
    Case 9
        
        FrameDevEfec.Visible = True
        Caption = "Listado"
        W = Me.FrameDevEfec.Width
        H = Me.FrameDevEfec.Height + 120
        
        
    Case 10
        
        FrameFormaPago.Visible = True
        Caption = "Listado"
        W = Me.FrameFormaPago.Width
        H = Me.FrameFormaPago.Height
    Case 11, 13, 43, 44
        
        FrameTransferencias.Visible = True
        
        If Opcion < 43 Then
            Label2(9).Caption = "Listado transferencias"
            SQL = "Listado trans."
            If Opcion = 11 Then
                'Puede ser transferencias o confirmings
                Caption = "PROVEEDORES"
            Else
                Caption = "ABONOS"
            End If
        
        Else
            SQL = "Listado "
            If Opcion = 43 Then
                'Puede ser transferencias o confirmings
                Caption = "Pagos domiciliados"
            Else
                Caption = "Caixa confirming"
            End If
            Label2(9).Caption = Caption
        End If
        Caption = SQL & " " & Caption
        W = Me.FrameTransferencias.Width
        H = Me.FrameTransferencias.Height + 60
        chkCartaAbonos.Visible = Opcion = 13
        
    Case 12
        
        FramePrevision.Visible = True
        Caption = "Listado"
        W = Me.FramePrevision.Width
        H = Me.FramePrevision.Height
        Text3(18).Text = Format(DateAdd("m", 2, Now), "dd/mm/yyyy")
        
        
    Case 15, 16, 17, 18, 33
        
        'Operaciones aseguradas
        '       Datos basicos
        '       Listado facturacion
        '       Impagados
        optAsegBasic(2).Visible = True 'Ordenar por poliza
        FrOrdenAseg1.Visible = True
        FrameASeg2.Visible = False
        FrameForpa.Visible = False
        FrameAsegAvisos.Visible = False
        Select Case Opcion
        Case 15
            '       Datos basicos
            SQL = "Fecha solicitud"
            Cad = "Datos b�sicos operaciones aseguradas"
            
        Case 16
            '       Listado facturacion
            SQL = "Fecha"
            Cad = "List. facturacion oper. aseguradas"
            FrOrdenAseg1.Visible = False
            FrameASeg2.Visible = True
            FrameForpa.Visible = True
        Case 17
            '       Listado impagados asegurados
            SQL = "Fecha aviso"
            Cad = "Impagados en operaciones aseguradas"
            
        Case 18
            optAsegBasic(2).Visible = False
            SQL = "Fecha vto"
            Cad = "Listado efectos operaciones aseguradas"
            
        Case 33
            FrameAsegAvisos.Visible = True
           ' FrOrdenAseg1.Visible = False
            SQL = "Fecha aviso falta pago"
            Cad = "Listados avisos aseguradoras"
            optAsegAvisos(0).Value = True
        End Select
        
        
        Label4(39).Caption = SQL
        Label2(11).Caption = Cad
        FrameAseg_Bas.Visible = True
        Caption = "Listado"
        W = Me.FrameAseg_Bas.Width
        H = Me.FrameAseg_Bas.Height
        
        
        
    Case 20
        H = FrameCobroGenerico.Height + 120
        W = FrameCobroGenerico.Width
        FrameCobroGenerico.Visible = True
        Caption = "Cuenta"
    Case 22
        
        
        For H = 0 To 1
            
            txtConcpto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 1)
            txtDescConcepto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 2)
        Next H
        Me.cboCompensaVto.Clear
        InsertaItemComboCompensaVto "No compensa sobre ning�n vencimiento", 0
        
        'Veremos si puede sobre un Vto o no
        H = RecuperaValor(CadenaDesdeOtroForm, 5)
        CONT = 0
        If H = 1 Then CONT = RecuperaValor(CadenaDesdeOtroForm, 6)
        FrameCambioFPCompensa.Visible = CONT > 0
        'chkCompensaVto.Value = 0
        'chkCompensaVto.Enabled = h = 1
        'chkCompensaVto.Caption = RecuperaValor(CadenaDesdeOtroForm, 6)
        CadenaDesdeOtroForm = ""
        H = FrameCompensaciones.Height + 120
        W = FrameCompensaciones.Width
        FrameCompensaciones.Visible = True
        Caption = "Compensacion efectos"
        Text3(23).Text = Format(Now, "dd/mm/yyyy")
        
        
    Case 23, 34
        '23.-  Contabilizar
        '34. Eliminar ya contabilizada
        
        
        
        
        'Tendremos el tipo de pago , talon o pagare
        Dim FP As Ctipoformapago
        Set FP = New Ctipoformapago
        
        If Opcion = 23 Then
            Label2(13).Caption = "Contabilizar recepci�n documentos"
            Caption = "Contabilizar"
        Else
            Label2(13).Caption = "Eliminar de recepci�n documentos"
            Caption = "Eliminar"
        End If
        
        'Cuenta beneficios gastos paras las diferencias si existieran
        'Si el total del talon es el total de las lineas entonces no mostrara los
        'datos del total. 0: igual   1  Mayor     2 Menor
        SQL = RecuperaValor(CadenaDesdeOtroForm, 2)
        I = CInt(SQL)
'        If CInt(SQL) > 0 Then
'            I = 1
'        Else
'            I = -1
'        End If
        
        Label4(55).Visible = I <> 0
        Image3(14).Visible = I <> 0
        txtCta(14).Visible = I <> 0
        DtxtCta(14).Visible = I <> 0
        Label6(28).Visible = I <> 0
        
        
        
        
        
        If I > 0 Then
            SQL = "Beneficios"
        Else
            SQL = "P�rdidas"
        End If
        
        If Opcion = 34 Then SQL = SQL & "(Deshacer apunte)"
        Label4(55).Caption = SQL

        


        '   No lleva ANALITICA
        If I <> 0 Then
            If Not vParam.autocoste Then I = 0
        End If
     
        Me.imgCCoste(0).Visible = I <> 0
        Me.txtCCost(0).Visible = I <> 0
        Label6(29).Visible = I <> 0
        Me.txtDescCCoste(0).Visible = I <> 0
     
        
        
        
        
        
        
        
        
        SQL = RecuperaValor(CadenaDesdeOtroForm, 1)
        I = CInt(SQL)
        If FP.Leer(I) = 0 Then
            If Opcion = 23 Then
                'Normal
                txtDiario(1).Text = FP.diaricli
                txtConcpto(2).Text = FP.condecli
                txtConcpto(3).Text = FP.conhacli
             Else
                'Eliminar. Iran cambiados
                txtDiario(1).Text = FP.diaricli
                txtConcpto(2).Text = FP.conhacli
                txtConcpto(3).Text = FP.condecli
                
                
             End If
                
            'Para que pinte la descripcion
            txtDiario_LostFocus 1
            txtConcpto_LostFocus 2
            txtConcpto_LostFocus 3
        End If
        
        
        
        
        H = 0
        If I = vbTalon Then
            SQL = "taloncta"
        Else
            SQL = "pagarecta"
        End If
        
        SQL = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
        If Len(SQL) = vEmpresa.DigitosUltimoNivel Then
            chkAgruparCtaPuente(0).Visible = True
            H = 1 '
        
            'Si esta configurado en parametrps, si la ultima vez lo marco seguira marcado
            If H = 1 Then H = CheckValueLeer("Agrup0")
            If H <> 1 Then H = 0
            chkAgruparCtaPuente(0).Value = H
            
        Else
            chkAgruparCtaPuente(0).Visible = False
        End If
        
        Set FP = Nothing
        
        If Label4(55).Visible Then '5055
            FrameRecepcionDocumentos.Height = 4815
            I = 4320
        Else
            FrameRecepcionDocumentos.Height = 3135
            I = 2640
        End If
        cmdRecepDocu.Top = I
        cmdCancelar(23).Top = I
        H = FrameRecepcionDocumentos.Height + 120
        W = FrameRecepcionDocumentos.Width
        FrameRecepcionDocumentos.Visible = True
        
        
            
        
        
        
    Case 24
        
        H = FrameListaRecep.Height + 120
        W = FrameListaRecep.Width
        FrameListaRecep.Visible = True
        
        
    Case 25
        
                
        H = frameListadoPagosBanco.Height + 120
        W = frameListadoPagosBanco.Width
        frameListadoPagosBanco.Visible = True
        
        
    Case 26
        'Si el total del talon es el total de las lineas entonces no mostrara los
        'datos del total. 0: igual   1  Mayor     2 Menor
        SQL = RecuperaValor(CadenaDesdeOtroForm, 1)
        If CCur(SQL) > 0 Then
            I = 1
        Else
            I = -1
        End If
        
        'Label4(55).Visible = True
        'Image3(14).Visible = True
        'txtCta(14).Visible = True
        'DtxtCta(14).Visible = I <> 0
        'Label6(28).Visible = I <> 0
        
         'If I > 0 Then
         '    SQL = "Beneficios"
         'Else
         '    SQL = "P�rdidas"
         'End If
         'Label4(55).Caption = SQL


        '   No lleva ANALITICA
        I = 1
        If Not vParam.autocoste Then I = 0
     
        Me.imgCCoste(0).Visible = I <> 0
        Me.txtCCost(0).Visible = I <> 0
        Label6(29).Visible = I <> 0
        Me.txtDescCCoste(0).Visible = I <> 0
        If I <> 0 Then Carga1ImagenAyuda imgCCoste(0), 1
        
        
'        h = FrameCancelRemTalPag.Height + 120
'        W = FrameCancelRemTalPag.Width
'        FrameCancelRemTalPag.Visible = True
        
        
    Case 27
                'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        H = FrameDividVto.Height + 120
        W = FrameDividVto.Width
        FrameDividVto.Visible = True
        
    Case 30
        H = FrameReclama.Height + 120
        W = FrameReclama.Width
        FrameReclama.Visible = True
        
    Case 31
        H = FrameGastosFijos.Height + 120
        W = FrameGastosFijos.Width
        FrameGastosFijos.Top = 0
        FrameGastosFijos.Left = 90
        FrameGastosFijos.Visible = True
        
    Case 35
        Me.txtVarios(2).Text = CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
        H = FrameGastosTranasferencia.Height + 120
        W = FrameGastosTranasferencia.Width
        FrameGastosTranasferencia.Visible = True
        
        
    Case 36
        
        
        H = FrameCompensaAbonosCliente.Height + 120
        W = FrameCompensaAbonosCliente.Width
        FrameCompensaAbonosCliente.Visible = True
        
        
        'cmdVtoDestino(1).Visible = (vUsu.Codigo Mod 100) = 0
        'Label1(1).Visible = (vUsu.Codigo Mod 100) = 0
        cmdVtoDestino(1).Visible = vUsu.Nivel = 0
        Label1(1).Visible = vUsu.Nivel = 0
        
        
    Case 38
        
        H = FrameRecaudaEjec.Height + 120
        W = FrameRecaudaEjec.Width
        FrameRecaudaEjec.Visible = True
        Fecha = DateAdd("yyyy", -4, Now)
        Text3(32).Text = Format(Fecha, "dd/mm/yyyy")
        
        
    Case 39, 40
        H = FrameOperAsegComunica.Height + 120
        W = FrameOperAsegComunica.Width
        FrameOperAsegComunica.Visible = True
        
        cargaEmpresasTesor ListView3
        Fecha = Now
        If Day(Now) < 15 Then Fecha = DateAdd("m", -1, Now)
        
        Text3(34).Text = "01/" & Format(Fecha, "mm/yyyy")
        
            
        Text3(35).Text = Format(Now, "dd/mm/yyyy")
        
        FrameSelEmpre1.BorderStyle = 0
        FrameFraPendOpAseg.BorderStyle = 0
        
        FrameSelEmpre1.Visible = Opcion = 39
        FrameFraPendOpAseg.Visible = Opcion = 40
        
        
        If Opcion = 39 Then
            Label2(22).Caption = "Comunicaci�n datos al seguro"
        Else
            Label2(22).Caption = "Fras. pendientes op. aseguradas"
        End If
        
        
    Case 42
        H = FrameNorma57Importar.Height + 120
        W = FrameNorma57Importar.Width
        FrameNorma57Importar.Visible = True
    
    
    Case 45
        
        H = FrameCobrosAgentesLin.Height + 120
        W = FrameCobrosAgentesLin.Width
        FrameCobrosAgentesLin.Visible = True
        Label3(50).Caption = ""
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    I = Opcion
    If Opcion = 13 Or I = 43 Or I = 44 Then I = 11
    
    'Aseguradas
    If Opcion >= 15 And Opcion <= 18 Then I = 15  'aseguradoas
    If Opcion = 33 Then I = 15 'aseguradoas
    If Opcion = 34 Then I = 23 'Eliminar recepcion documento
    If Opcion = 40 Then I = 39
    Me.cmdCancelar(I).Cancel = True
    
    PonerFrameProgreso

End Sub


Private Sub PonerFrameProgreso()
Dim I As Integer

    'Ponemos el frame al pricnipio de todo
    FrameProgreso.Visible = False
    FrameProgreso.ZOrder 0
    
    'lo ubicamos
    'Posicion horizintal WIDTH
    I = Me.Width - FrameProgreso.Width
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Left = I
    'Posicion  VERTICAL HEIGHT
    I = Me.Height - FrameProgreso.Height
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Top = I
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 1 Then
        CheckValueGuardar "Listcta", CByte(Abs(Me.optCuenta(0).Value))
        CheckValueGuardar "Infapa", chkApaisado(0)
    End If
    If Opcion = 23 Then CheckValueGuardar "Agrup0", Me.chkAgruparCtaPuente(0)
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtAgente(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescAgente(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    DevfrmCCtas = CadenaDevuelta
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub







Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    'Si no habia cuenta
    
        txtCta(0).Text = RecuperaValor(CadenaSeleccion, 1)
        DtxtCta(0).Text = RecuperaValor(CadenaSeleccion, 2)
        txtCta(1).Text = RecuperaValor(CadenaSeleccion, 1)
        DtxtCta(1).Text = RecuperaValor(CadenaSeleccion, 2)
    
    'El dpto
    txtDpto(RC).Text = RecuperaValor(CadenaSeleccion, 3)
    txtDescDpto(RC).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtFPago(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescFPago(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    txtSerie(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Index = 17 Then PonerVtosCompensacionCliente
End Sub










Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
    Select Case Index
    Case 0
            C = "Compensaciones" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Cuando compense sobre un vencimiento al marcar la opci�n " & vbCrLf
            C = C & Space(10) & Me.chkCompensa.Caption & vbCrLf
            C = C & "se modificar� el importe vencimiento poniendo el total a compensar  y en importe cobrado un cero"
    Case 1
            C = "Asegurados" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Fecha 'hasta' es campo obligado para considerar la fecha de baja de los asegurados." & vbCrLf
            C = C & "En los listados saldr�n aquellos que si tienen fecha de baja , es superior al hasta solicitado "
            
            'ALZIRA
            C = C & vbCrLf & vbCrLf & "Comunicaci�n datos seguro" & vbCrLf & "Salen TODAS las facturas entre el periodo seleccionado para los"
            C = C & " clientes asegurados"
            
    End Select
    MsgBox C, vbInformation

End Sub

Private Sub Imagente_Click(Index As Integer)
    Set frmA = New frmAgentes
    RC = Index
    frmA.DatosADevolverBusqueda = "0|1|"
    frmA.Show vbModal
    Set frmA = Nothing
End Sub

Private Sub ImageSe_Click(Index As Integer)
    RC = Index
    Set frmS = New frmSerie
    frmS.DatosADevolverBusqueda = "0"
    frmS.Show vbModal
    Set frmS = Nothing
End Sub

Private Sub imgCarta_Click()
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    DevfrmCCtas = ""
    frmB.vSQL = ""
    
    '###A mano
    frmB.vDevuelve = "0|1|"   'Siempre el 0
    
    frmB.vSelElem = 1

    Cad = "Codigo|codcarta|N|15�"
    Cad = Cad & "Descripcion|descarta|T|65�"
    frmB.vCampos = Cad
    frmB.vTabla = "scartas"
    frmB.vTitulo = "Cartas reclamaci�n"
    frmB.Show vbModal
    Set frmB = Nothing
    If DevfrmCCtas <> "" Then
        Me.txtDescCarta.Text = RecuperaValor(DevfrmCCtas, 2)
        txtCarta.Text = RecuperaValor(DevfrmCCtas, 1)
    End If
End Sub

Private Sub imgCCoste_Click(Index As Integer)
    LanzaBuscaGrid Index, 2
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For I = 1 To Me.ListView3.ListItems.Count
        Me.ListView3.ListItems(I).Checked = (Index = 1)
    Next
        
End Sub

Private Sub imgConcepto_Click(Index As Integer)
    LanzaBuscaGrid Index, 1
End Sub

Private Sub imgCtaBanc_Click(Index As Integer)
    SQL = ""
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
    If SQL <> "" Then
        txtCtaBanc(Index).Text = RecuperaValor(SQL, 1)
        Me.txtDescBanc(Index).Text = RecuperaValor(SQL, 2)
    End If
End Sub

Private Sub imgDiario_Click(Index As Integer)
    LanzaBuscaGrid Index, 0
End Sub

Private Sub imgDpto_Click(Index As Integer)
    SQL = "NO"
    If txtCta(1).Text <> "" And txtCta(0).Text <> "" Then
        
        If txtCta(1).Text <> txtCta(0).Text Then
            MsgBox "Debe seleccionar un mismo cliente", vbExclamation
            txtDpto(Index).Text = ""
            SQL = ""
        End If
    End If
    If SQL = "" Then Exit Sub
        
    Set frmD = New frmDepartamentos
    RC = Index
    frmD.DatosADevolverBusqueda = "1|2|"
    frmD.vCuenta = txtCta(0).Text
    frmD.Show vbModal
    Set frmD = Nothing
End Sub


Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    'Set frmCta = New frmColCtas
    Set frmP = New frmFormaPago
    RC = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
End Sub






Private Sub imgGastoFijo_Click(Index As Integer)
     LanzaBuscaGrid Index, 3
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lwCompenCli_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    If Trim(Item.SubItems(6)) = "" Then
        'Es un abono
        Cobro = False
        C = -C
    
    End If
    
    'Si no es checkear cambiamos los signos
    If Not Item.Checked Then C = -C
    
    I = 0
    If Not Cobro Then I = 1
    
    Me.txtimpNoEdit(I).Tag = Me.txtimpNoEdit(I).Tag + C
    txtimpNoEdit(I).Text = Format(Abs(txtimpNoEdit(I).Tag))
    txtimpNoEdit(2).Text = Format(CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag), FormatoImporte)
            
End Sub

Private Sub optAsegAvisos_Click(Index As Integer)
    If Index = 0 Then
        Label4(39).Caption = "Fecha aviso falta pago"
    ElseIf Index = 1 Then
        Label4(39).Caption = "Fecha aviso prorroga"
    Else
        Label4(39).Caption = "Fecha aviso siniestro"
    End If
    
End Sub

Private Sub optAsegAvisos_KeyPress(Index As Integer, KeyAscii As Integer)
     KeyPressGral KeyAscii
End Sub

Private Sub optAsegBasic_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub optCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub


Private Sub optImpago_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub optLCobros_Click(Index As Integer)
    Me.Check1.Enabled = Me.optLCobros(1).Value
    Me.Check2.Enabled = Not Me.Check1.Enabled
End Sub

Private Sub optLCobros_KeyPress(Index As Integer, KeyAscii As Integer)
KeyPressGral KeyAscii
End Sub

Private Sub optPrevision_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub optProv_Click(Index As Integer)
    Me.chkProv.Enabled = Me.optProv(1).Value
    Me.chkProv2.Enabled = Not Me.chkProv.Enabled
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub















Private Sub txtAgente_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtAgente_LostFocus(Index As Integer)

    SQL = ""
    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    If txtAgente(Index).Text <> "" Then
        
        If Not IsNumeric(txtAgente(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            txtAgente(Index).Text = ""
            SubSetFocus txtAgente(Index)
        Else
            txtAgente(Index).Text = Val(txtAgente(Index).Text)
            SQL = DevuelveDesdeBD("nombre", "agentes", "codigo", txtAgente(Index).Text, "N")
            If SQL = "" Then SQL = "AGENTE NO ENCONTRADO"
        End If
    End If
    Me.txtDescAgente(Index).Text = SQL
        
End Sub





Private Sub txtCarta_GotFocus()
    PonFoco txtCarta
End Sub

Private Sub txtCarta_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtCarta_LostFocus()
    SQL = ""
    txtCarta.Text = Trim(txtCarta.Text)
    If txtCarta.Text <> "" Then
        
        If Not IsNumeric(txtCarta.Text) Then
            MsgBox "Campo num�rico", vbExclamation
            txtCarta.Text = ""
            SubSetFocus txtCarta
        Else
            txtCarta.Text = Val(txtCarta.Text)
            SQL = DevuelveDesdeBD("descarta", "scartas", "codcarta", txtCarta.Text, "N")
            If SQL = "" Then txtCarta.Text = ""
        End If
    End If
    Me.txtDescCarta.Text = SQL
End Sub





Private Sub txtCCost_GotFocus(Index As Integer)
    PonFoco txtConcpto(Index)
End Sub

Private Sub txtCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtCCost_LostFocus(Index As Integer)
    SQL = ""
    txtCCost(Index).Text = Trim(txtCCost(Index).Text)
    If txtCCost(Index).Text <> "" Then
        

            txtConcpto(Index).Text = Val(txtConcpto(Index).Text)
            SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtCCost(Index).Text, "T")
            If SQL = "" Then
                MsgBox "No existe el centro de coste: " & Me.txtCCost(Index).Text, vbExclamation
                Me.txtCCost(Index).Text = ""
            End If
        If txtCCost(Index).Text = "" Then SubSetFocus txtCCost(Index)
    End If
    Me.txtDescCCoste(Index).Text = SQL
End Sub

Private Sub txtConcpto_GotFocus(Index As Integer)
     PonFoco txtConcpto(Index)
End Sub

Private Sub txtConcpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtConcpto_LostFocus(Index As Integer)
    SQL = ""
    txtConcpto(Index).Text = Trim(txtConcpto(Index).Text)
    If txtConcpto(Index).Text <> "" Then
        
        If Not IsNumeric(txtConcpto(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            txtConcpto(Index).Text = ""
        Else
            txtConcpto(Index).Text = Val(txtConcpto(Index).Text)
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcpto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "No existe el concepto: " & Me.txtConcpto(Index).Text, vbExclamation
                Me.txtConcpto(Index).Text = ""
            End If
        End If
        If txtConcpto(Index).Text = "" Then SubSetFocus txtConcpto(Index)
    End If
    Me.txtDescConcepto(Index).Text = SQL
    
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    If Index = 6 Then
        'NO se ha cambiado nada de la cuenta
        If txtCta(6).Text = txtCta(6).Tag Then
        
            Exit Sub
        Else
            txtDpto(0).Text = ""
            txtDpto(1).Text = ""
            txtDescDpto(0).Text = ""
            txtDescDpto(0).Text = ""
        End If
    End If
     
     
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
       ' txtCta(6).Tag = txtCta(6).Text
        Exit Sub
    End If
    
    If Index = 6 Then
        If txtCta(0).Text <> "" Or txtCta(1).Text <> "" Then
            MsgBox "Si selecciona desde / hasta cliente no podra seleccionar departamento", vbExclamation
            txtCta(6).Text = ""
            txtCta(6).Tag = txtCta(6).Text
            Exit Sub
        End If
        
    Else
        If Index = 0 Or Index = 1 Then
            If txtCta(6).Text <> "" Then
                MsgBox "Si seleciona departamento no puede seleccionar desde / hasta  cliente", vbExclamation
                txtCta(Index).Text = ""
                txtCta(6).Tag = txtCta(6).Text
                Exit Sub
            End If
        End If
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        MsgBox "La cuenta debe ser num�rica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        txtCta(6).Tag = txtCta(6).Text
        Ponerfoco txtCta(Index)
        
        If Index = 17 Then PonerVtosCompensacionCliente
        
        Exit Sub
    End If
    
    Select Case Index
    Case 0 To 7, 11, 12, 15, 16, 18, 19
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = SQL
            End If
            
            
            'Index=1. Cliente en listado de cobros. Si pongo el desde pongo el hasta lo mismo
            If Index = 1 Then
                
                If Len(Cta) = vEmpresa.DigitosUltimoNivel Then
                    txtCta(0).Text = Cta
                    DtxtCta(0).Text = DtxtCta(1).Text
                End If
            End If
            
        End If
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, SQL) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            
            
        Else
            MsgBox SQL, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
        If Index = 17 Then PonerVtosCompensacionCliente
        
    End Select
    txtCta(6).Tag = txtCta(6).Text
End Sub







Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    
    SQL = ""
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    If txtDiario(Index).Text <> "" Then
        
        If Not IsNumeric(txtDiario(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            txtDiario(Index).Text = ""
            SubSetFocus txtDiario(Index)
        Else
            txtDiario(Index).Text = Val(txtDiario(Index).Text)
            SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text, "N")
            
            If SQL = "" Then
                MsgBox "No existe el diario: " & Me.txtDiario(Index).Text, vbExclamation
                Me.txtDiario(Index).Text = ""
                Ponerfoco txtDiario(Index)
            End If
        End If
    End If
    Me.txtDescDiario(Index).Text = SQL
     
End Sub





Private Sub txtGastoFijo_GotFocus(Index As Integer)
    PonFoco txtGastoFijo(Index)
End Sub

Private Sub txtGastoFijo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtGastoFijo_LostFocus(Index As Integer)
    
    SQL = ""
    txtGastoFijo(Index).Text = Trim(txtGastoFijo(Index).Text)
    If txtGastoFijo(Index).Text <> "" Then
        
        If Not IsNumeric(txtGastoFijo(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            txtGastoFijo(Index).Text = ""
            SubSetFocus txtGastoFijo(Index)
        Else
            'sgastfij codigo Descripcion
            txtGastoFijo(Index).Text = Val(txtGastoFijo(Index).Text)
            SQL = DevuelveDesdeBD("Descripcion", "sgastfij", "codigo", txtGastoFijo(Index).Text, "N")
            
            If SQL = "" Then
                MsgBox "No existe el gasto fijo: " & Me.txtGastoFijo(Index).Text, vbExclamation

            End If
        End If
    End If
    Me.txtDescGastoFijo(Index).Text = SQL
     
End Sub











Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    If txtImporte(Index).Text = "" Then Exit Sub
    Mal = False
    If Not EsNumerico(txtImporte(Index).Text) Then Mal = True

    If Not Mal Then Mal = Not CadenaCurrency(txtImporte(Index).Text, Importe)

    If Mal Then
        txtImporte(Index).Text = ""
        txtImporte(Index).SetFocus

    Else
        txtImporte(Index).Text = Format(Importe, FormatoImporte)
    End If
    




End Sub









'Private Sub PonerNiveles()
'Dim i As Integer
'Dim J As Integer
'
'
'
'
'    Check1(10).Visible = True
'    For i = 1 To vEmpresa.numnivel - 1
'        J = DigitosNivel(i)
'        cad = "Digitos: " & J
'        Check1(i).Visible = True
'        Me.Check1(i).Caption = cad
'
'        'Para los de balance presupuestario
'        Me.ChkCtaPre(i).Visible = True
'        Me.ChkCtaPre(i).Caption = cad
'        'para los de resumen dairio
'        Me.ChkNivelRes(i).Visible = True
'        Me.ChkNivelRes(i).Caption = cad
'
'        'Consolidado
'        Me.ChkConso(i).Visible = True
'        Me.ChkConso(i).Caption = cad
'
'        chkcmp(i).Caption = cad
'        chkcmp(i).Visible = True
'
'        Combo2.AddItem "Nivel :   " & i
'        Combo2.ItemData(Combo2.NewIndex) = J
'    Next i
'    For i = vEmpresa.numnivel To 9
'        Check1(i).Visible = False
'        Me.ChkCtaPre(i).Visible = False
'        Me.ChkNivelRes(i).Visible = False
'        chkcmp(i).Visible = False
'        ChkConso(i).Visible = False
'    Next i
'
'End Sub






Private Sub CargarComboFecha()
'Dim J As Integer
'
'
'QueCombosFechaCargar "0|1|2|"
'
'
''Y ademas deshabilitamos los niveles no utilizados por la aplicacion
'For i = vEmpresa.numnivel To 9
'    Check2(i).Visible = False
'    Me.chkCtaExplo(i).Visible = False
'    chkCtaExploC(i).Visible = False
'    chkAce(i).Visible = False
'Next i
'
'For i = 1 To vEmpresa.numnivel - 1
'    J = DigitosNivel(i)
'    Check2(i).Visible = True
'    Check2(i).Caption = "Digitos: " & J
'    chkCtaExplo(i).Visible = True
'    chkCtaExplo(i).Caption = "Digitos: " & J
'    chkAce(i).Visible = True
'    chkAce(i).Caption = "Digitos: " & J
'    chkCtaExploC(i).Visible = True
'    chkCtaExploC(i).Caption = "Digitos: " & J
'Next i
'
'
'
'
''Cargamos le combo de resalte de fechas
'Combo3.AddItem "Sin remarcar"
'Combo3.ItemData(Combo3.NewIndex) = 1000
'For i = 1 To vEmpresa.numnivel - 1
'    Combo3.AddItem "Nivel " & i
'    Combo3.ItemData(Combo3.NewIndex) = i
'Next i
End Sub




























Private Sub QueCombosFechaCargar(Lista As String)
'Dim L As Integer
'
'L = 1
'Do
'    cad = RecuperaValor(Lista, L)
'    If cad <> "" Then
'        i = Val(cad)
'        With cmbFecha(i)
'            .Clear
'            For Cont = 1 To 12
'                RC = "25/" & Cont & "/2002"
'                RC = Format(RC, "mmmm") 'Devuelve el mes
'                .AddItem RC
'            Next Cont
'        End With
'    End If
'    L = L + 1
'Loop Until cad = ""
End Sub











Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function





Private Sub txtCtaBanc_GotFocus(Index As Integer)
    PonFoco txtCtaBanc(Index)
End Sub

Private Sub txtCtaBanc_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtCtaBanc_LostFocus(Index As Integer)
    txtCtaBanc(Index).Text = Trim(txtCtaBanc(Index).Text)
    If txtCtaBanc(Index).Text = "" Then
        txtDescBanc(Index).Text = ""
        Exit Sub
    End If
    
    Cad = txtCtaBanc(Index).Text
    I = CuentaCorrectaUltimoNivelSIN(Cad, SQL)
    If I = 0 Then
        MsgBox "NO existe la cuenta: " & txtCtaBanc(Index).Text, vbExclamation
        SQL = ""
        Cad = ""
    Else
        Cad = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", Cad, "T")
        If Cad = "" Then
            MsgBox "Cuenta no asoaciada a ningun banco", vbExclamation
            SQL = ""
            I = 0
        End If
    End If
    
    txtCtaBanc(Index).Text = Cad
    Me.txtDescBanc(Index).Text = SQL
    If I = 0 Then Ponerfoco txtCtaBanc(Index)
    
End Sub

Private Sub txtDias_GotFocus()
    PonFoco txtDias
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtDias_LostFocus()
    txtDias.Text = Trim(txtDias.Text)
    If txtDias.Text <> "" Then
        If Not IsNumeric(txtDias.Text) Then
            MsgBox "Numero de dias debe ser num�rico", vbExclamation
            txtDias.Text = ""
            SubSetFocus txtDias
        End If
    End If
End Sub



Private Sub txtDpto_GotFocus(Index As Integer)
    PonFoco txtDpto(Index)
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
    
    'Pierde foco
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If txtDpto(Index).Text = "" Then
        Me.txtDescDpto(Index).Text = ""
        Exit Sub
    End If
    
    SQL = "NO"
    If txtCta(1).Text = "" Or txtCta(0).Text = "" Then
        MsgBox "Debe seleccionar un unico cliente", vbExclamation
        txtDpto(Index).Text = ""
        SQL = ""
    Else
        If txtCta(1).Text <> txtCta(0).Text Then
            MsgBox "Debe seleccionar un mismo cliente", vbExclamation
            txtDpto(Index).Text = ""
            SQL = ""
        End If
    End If
    
    If SQL <> "" Then
        SQL = ""
        If txtCta(1).Text <> "" Then
            If txtDpto(Index).Text <> "" Then
                If Not IsNumeric(txtDpto(Index).Text) Then
                      MsgBox "Codigo departamento debe ser numerico: " & txtDpto(Index).Text
                      txtDpto(Index).Text = ""
                Else
                      'Comproamos en la BD
                       Set RS = New ADODB.Recordset
                       Cad = "Select descripcion from departamentos where codmacta='" & txtCta(0).Text
                       Cad = Cad & "' AND Dpto = " & txtDpto(Index).Text
                       RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                       If Not RS.EOF Then SQL = DBLet(RS.Fields(0), "T")
                       RS.Close
                       Set RS = Nothing
                End If
            End If
        Else
            If txtDpto(Index).Text <> "" Then
                MsgBox "Seleccione un cliente", vbExclamation
                txtDpto(Index).Text = ""
            End If
        End If
    End If
    Me.txtDescDpto(Index).Text = SQL
End Sub

Private Sub txtFPago_GotFocus(Index As Integer)
    PonFoco txtFPago(Index)
End Sub

Private Sub txtFPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub


Private Sub txtFPago_LostFocus(Index As Integer)
    If ComprobarCampoENlazado(txtFPago(Index), txtDescFPago(Index), "N") > 0 Then
        If txtFPago(Index).Text <> "" Then
            'Tiene valor.
            SQL = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", txtFPago(Index).Text, "N")
            If SQL = "" Then SQL = "Codigo no encontrado"
            txtDescFPago(Index).Text = SQL
        Else
            'Era un error
            SubSetFocus txtFPago(Index)
        End If
    End If
End Sub




Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If

    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function



Private Function DesdeHasta(Tipo As String, Desde As Integer, Hasta As Integer, Optional Des As String)
Dim C As String
    C = ""
    Select Case Tipo
    Case "F"
        'Son los text3(desde)....
        If Text3(Desde).Text <> "" Then C = C & "Desde " & Text3(Desde).Text
        
        If Text3(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & Text3(Hasta).Text
        End If
        
    Case "C"
        'Cuenta
        If txtCta(Desde).Text <> "" Then C = "Desde " & txtCta(Desde).Text & "-" & DtxtCta(Desde).Text
        
        
        If txtCta(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCta(Hasta).Text & "-" & DtxtCta(Hasta).Text
        End If
        
        
    Case "FP"
        'FORMA DE PAGO
        If txtFPago(Desde).Text <> "" Then C = "Desde " & txtFPago(Desde).Text & "-" & txtDescFPago(Desde).Text
        
        
        If txtFPago(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtFPago(Hasta).Text & "-" & txtDescFPago(Hasta).Text
        End If
    
    Case "BANCO"
        '    'txtCtaBanc  txtDescBanc
        If txtCtaBanc(Desde).Text <> "" Then C = "Desde " & txtCtaBanc(Desde).Text & "-" & txtDescBanc(Desde).Text
        
        If txtCtaBanc(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCtaBanc(Hasta).Text & "-" & txtDescBanc(Hasta).Text
        End If
    
    
    Case "S"
        'Serie
        If txtSerie(Desde).Text <> "" Then C = C & "Desde " & txtSerie(Desde).Text
        
        If txtSerie(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtSerie(Hasta).Text
        End If
    
    Case "NF"
        'Numero factura
        If txtNumfac(Desde).Text <> "" Then C = C & "Desde " & txtNumfac(Desde).Text
        
        If txtNumfac(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtNumfac(Hasta).Text
        End If
    
    Case "GF"
        'Gasto fijo
        
        If txtGastoFijo(Desde).Text <> "" Then C = C & "Desde " & txtGastoFijo(Desde).Text & " - " & Me.txtDescGastoFijo(Desde).Text
        
        If txtGastoFijo(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtGastoFijo(Hasta).Text & " - " & Me.txtDescGastoFijo(Hasta).Text
        End If
    
    
    End Select
    If C <> "" Then C = "  " & Des & " " & C
    DesdeHasta = C
End Function


Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub





'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Private Function CobrosPendientesCliente(ByVal ListadoCuentas As String) As Boolean
Dim TieneRemesa As Boolean
Dim RemesaTalones As Boolean
Dim RemesaPagares As Boolean
Dim RemesaEfectos As Boolean
Dim SePuedeRemesar As Boolean
Dim InsertarLinea As Boolean


Dim CadenaInsert As String

    On Error GoTo ECobrosPendientesCliente
    CobrosPendientesCliente = False

    
    'De parametros contapagarepte contatalonpte
    Cad = DevuelveDesdeBD("contatalonpte", "paramtesor", "codigo", "1")
    If Cad = "" Then Cad = "0"
    RemesaTalones = (Val(Cad) = 1)
    
    Cad = DevuelveDesdeBD("contapagarepte", "paramtesor", "codigo", "1")
    If Cad = "" Then Cad = "0"
    RemesaPagares = (Val(Cad) = 1)
    
    Cad = DevuelveDesdeBD("contaefecpte", "paramtesor", "codigo", "1")
    If Cad = "" Then Cad = "0"
    RemesaEfectos = (Val(Cad) = 1)
    

    
    
    
    
    'Trozo basico
    Cad = " FROM scobro ,cuentas,sforpa ,stipoformapago"
    Cad = Cad & " WHERE  scobro.codmacta = cuentas.codmacta"
    Cad = Cad & " AND scobro.codforpa = sforpa.codforpa"
    Cad = Cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    
    'Fecha factura
    RC = CampoABD(Text3(1), "F", "fecfaccl", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Text3(2), "F", "fecfaccl", False)
    If RC <> "" Then Cad = Cad & " AND " & RC



    'Se me habia olvidado
    'A G E N T E
    RC = CampoABD(txtAgente(0), "N", "agente", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtAgente(1), "N", "agente", False)
    If RC <> "" Then Cad = Cad & " AND " & RC




    'Fecha vencimiento
    RC = CampoABD(Text3(19), "F", "fecvenci", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Text3(20), "F", "fecvenci", False)
    If RC <> "" Then Cad = Cad & " AND " & RC

    'SERIE
    RC = CampoABD(txtSerie(0), "T", "numserie", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtSerie(1), "T", "numserie", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    'Numero factura
    RC = CampoABD(txtNumfac(0), "T", "codfaccl", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtNumfac(1), "T", "codfaccl", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    


    'Cliente
    RC = CampoABD(txtCta(1), "T", "scobro.codmacta", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtCta(0), "T", "scobro.codmacta", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    'Forma PAGO
    RC = CampoABD(txtFPago(0), "T", "scobro.codforpa", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtFPago(1), "T", "scobro.codforpa", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    'Cliente con departamento
    'If txtCta(0).Text <> "" Then
    '    If cad <> "" Then cad = cad & " AND "
    '    cad = cad & " scobro.codmacta = '" & txtCta(6).Text & "'"
    'End If
    
    'Departamento
    RC = CampoABD(Me.txtDpto(0), "N", "departamento", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Me.txtDpto(1), "N", "departamento", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    'Solo los NO remesar
    If Me.chkNOremesar.Value = 1 Then
        Cad = Cad & " AND noremesar = 1 "
    End If
    
    'Docuemtno recibido y devuelto. Los combos  recedocu  Devuelto
    If Me.cboCobro(0).ListIndex > 0 Then Cad = Cad & " AND recedocu = " & cboCobro(0).ItemData(cboCobro(0).ListIndex)
    If Me.cboCobro(1).ListIndex > 0 Then Cad = Cad & " AND Devuelto = " & cboCobro(1).ItemData(cboCobro(1).ListIndex)
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        SQL = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then SQL = SQL & ","
                NumRegElim = 2
                SQL = SQL & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        Cad = Cad & " AND scobro.codmacta IN (" & SQL & ")"
    End If
    
    
    
    'Si ha marcado alguna forma de pago
    RC = PonerTipoPagoCobro_(True, False)
    If RC <> "" Then Cad = Cad & " AND tipoformapago IN " & RC
    RC = ""
    
    'Contador
    SQL = "Select count(*) "
    SQL = SQL & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not RS.EOF Then
        'Total registros
        TotalRegistros = RS.Fields(0)
    End If
    RS.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    'Marzo 2015
    'Si agrupamos por forma de pago, agruparemos antes por Tipo de pago
    If chkFormaPago.Value = 1 Then Conn.Execute "DELETE FROM Usuarios.zsimulainm where codusu = " & vUsu.Codigo
    
    
    
    
    
    SQL = "SELECT scobro.*, cuentas.nommacta, nifdatos,stipoformapago.descformapago ,stipoformapago.tipoformapago,nomforpa " & Cad
    ' ----------------
    '  20 Diciembre 2005
    '  Remesados o no remesados
    '
    CONT = 1
    If Me.ChkAgruparSituacion.Value = 1 Then
        '
        CONT = 0
    Else
        If Me.chkEfectosContabilizados.Value = 1 Then CONT = 0
    End If
    If CONT = 1 Then SQL = SQL & " AND codrem is null"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    TieneRemesa = False
    SQL = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    SQL = SQL & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,gastos,vencido,Situacion"
    'Nuevo Enero 2009
    'Si esta apaisado ponemos los departamentos
    If Me.chkApaisado(0).Value = 1 Then
        SQL = SQL & ",coddirec,nomdirec"
    Else
        'Metemos el NIF para futors listados. Pej. El de cobors por cliente lo pondra
        SQL = SQL & ",nomdirec"
    End If
    SQL = SQL & ",devuelto,recibido"
    'SQL = SQL & ",observa) VALUES (" & vUsu.Codigo & ",'"
    'Dic 2013 . Acelerar proceso
    SQL = SQL & ",observa) VALUES "
    
    
    CadenaInsert = "" 'acelerar carga datos
    Fecha = CDate(Text3(0).Text)
    While Not RS.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
        
        'If Rs!codmacta = "4300019" Then Stop
        
        Cad = RS!NUmSerie & "','" & Format(RS!codfaccl, "0000") & "','" & Format(RS!fecfaccl, FormatoFecha) & "'," & RS!numorden
        
        'Modificacion. Enero 2010. Tiene k aparacer la forma de pago, no el tipo
        'Cad = Cad & "," & Rs!codforpa & ",'" & DevNombreSQL(Rs!descformapago) & "','"
        Cad = Cad & "," & RS!codforpa & ",'" & DevNombreSQL(RS!nomforpa) & "','"
        
        Cad = Cad & RS!codmacta & "','" & DevNombreSQL(RS!Nommacta) & "','"
        Cad = Cad & Format(RS!fecvenci, FormatoFecha) & "',"
        Cad = Cad & TransformaComasPuntos(CStr(RS!impvenci)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(RS!impcobro) Then
            Cad = Cad & TransformaComasPuntos(CStr(RS!impcobro))
        Else
            Cad = Cad & "0"
        End If
        
        'Gastos
        Cad = Cad & "," & TransformaComasPuntos(DBLet(RS!Gastos, "N"))
        
        If Fecha > RS!fecvenci Then
            Cad = Cad & ",1"
        Else
            Cad = Cad & ",0"
        End If

        'Hay que a�adir la situacion. Bien sea juridica....
        ' Si NO agrupa por situacion, en ese campo metere la referencia del cobro (rs!referencia)
         'vbTalon = 2 vbPagare = 3
        InsertarLinea = True
        
        If Me.ChkAgruparSituacion.Value = 0 Then
            Cad = Cad & ",'" & DevNombreSQL(DBLet(RS!referencia, "T")) & "'"
            
            'Si no agrupa por situacion de vto y no tiene el riesgo marcado
            'Talones pagares
            'Si se ha recepcionado documento, NO deben salir
            
            
            'Enero 2010
            'Comentamos esto ya que tiene la marca de recibidos si/no
'            If Me.chkEfectosContabilizados.Value = 0 Then
'                If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
'                    If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
'                End If
'            End If

            
        Else
            'La situacion.
            'Si esta en situacion juridica.
            ' Si no, si esta devuelto y no es una remesa
            ' y luego si es una remesa, sitaucion o devuelto
            If RS!situacionjuri = 1 Then
                Cad = Cad & ",'SITUACION JURIDICA'"
            Else
                'Cambio Marzo 2009
                ' Ahora tb se remesan los pagares y talones
                
                If Not IsNull(RS!siturem) Then
                    TieneRemesa = True
                    Cad = Cad & ",'R" & Format(RS!AnyoRem, "0000") & Format(RS!CodRem, "0000000000") & "'"
                    
                Else
                    
                    If RS!Devuelto = 1 Then
                        Cad = Cad & ",'DEVUELTO'"
                    Else
                            
                        SePuedeRemesar = False
                        If RemesaEfectos Then SePuedeRemesar = RS!tipoformapago = vbTipoPagoRemesa
                        If RemesaPagares Then SePuedeRemesar = RS!tipoformapago = vbPagare
                        If RemesaTalones Then SePuedeRemesar = RS!tipoformapago = vbTalon
                        
                    
                        If Not SePuedeRemesar Then
                            Cad = Cad & ",'PENDIENTE COBRO'"
                        Else
                            Cad = Cad & ",'PENDIENTE REMESAR'" '& Rs!anyorem
                        End If
                        
                        
                        
                        'Talones pagares
                        'Si se ha recepcionado documento, NO deben salir
                        'ENERO 2010
                        'Tiene la marca de recibidos
                        
                        'If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
                        '    If Me.chkEfectosContabilizados.Value = 0 Then
                        '        If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
                        '    End If
                        'End If
                        
                    
                    End If  'De devuelto
               End If  'de SITUREM null
            End If ' de situacion juridica
        End If  'de agrupa por sitacuib
        Cad = Cad & ","
        If Me.chkApaisado(0).Value = 1 Then
            'SI carga departamentos. Esto podriamos mejorar la velocidad si
            'pregarmos un rs o en la select linkamos con departamento...
            If IsNull(RS!departamento) Then
                Cad = Cad & "NULL,NULL,"
            Else
                Cad = Cad & "'" & RS!departamento & "','"
                Cad = Cad & DevNombreSQL(DevuelveDesdeBD("Descripcion", "departamentos", "codmacta = '" & RS!codmacta & "' AND dpto", RS!departamento, "N")) & "',"
            End If
            
        Else
            'Nif datos
            'Stop
             Cad = Cad & "'" & DevNombreSQL(DBLet(RS!nifdatos, "T")) & "',"
        End If
        
        If DBLet(RS!Devuelto, "N") = 0 Then
            Cad = Cad & "'',"
        Else
            Cad = Cad & "'S',"
        End If
        If DBLet(RS!recedocu, "N") = 0 Then
            Cad = Cad & "''"
        Else
            Cad = Cad & "'S'"
        End If
            
        Cad = Cad & ",'"
        If Me.ChkObserva.Value Then
            Cad = Cad & DevNombreSQL(DBLet(RS!obs, "T"))
'        Else
'            Cad = Cad & "''"
        End If
        Cad = Cad & "')"
        
        If InsertarLinea Then
        
            CadenaInsert = CadenaInsert & ", (" & vUsu.Codigo & ",'" & Cad
        
            If Len(CadenaInsert) > 20000 Then
                Cad = SQL & Mid(CadenaInsert, 2)
                Conn.Execute Cad
                CadenaInsert = ""
            End If
            'Cad = SQL & Cad
            'Conn.Execute Cad
        Else
            'Tiramos para atras el contador total
            CONT = CONT - 1
        End If
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
    If Len(CadenaInsert) > 0 Then
        Cad = SQL & Mid(CadenaInsert, 2)
        Conn.Execute Cad
        CadenaInsert = ""
    End If

    
    'Si esta seleccacona SITIACUIN VENCIMIENTO
    ' y tenia remesas , entonces updateo la tabla poniendo
    ' la situacion de la remesa
    If TieneRemesa Then
        Cad = "Select codigo,anyo,  descsituacion"
        Cad = Cad & " from remesas left join tiposituacionrem on situacion=situacio"
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Debug.Print RS!Codigo
            If Not IsNull(RS!descsituacion) Then
                Cad = "R" & Format(RS!Anyo, "0000") & Format(RS!Codigo, "0000000000")
                Cad = " WHERE situacion='" & Cad & "'"
                Cad = "UPDATE Usuarios.zpendientes set Situacion='Remesados: " & RS!descsituacion & "' " & Cad
                Conn.Execute Cad
            End If
            RS.MoveNext
        Wend
        RS.Close
    End If
    
    'Marzo 2015.
    'Nivel de anidacion para los agrupados por forma de pago
    ' que es TIPO DE PAGO
    If chkFormaPago.Value = 1 Then
    
        Cad = "select codforpa from Usuarios.zpendientes where codusu =" & vUsu.Codigo & " group by 1"
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RS.EOF
            Cad = Cad & ", " & RS!codforpa
            RS.MoveNext
        Wend
        RS.Close
        
        If Cad <> "" Then
            Cad = Mid(Cad, 2)
            Cad = " and codforpa IN (" & Cad & ")"
            Cad = " FROM sforpa , stipoformapago WHERE sforpa.tipforpa=stipoformapago.tipoformapago" & Cad
            Cad = "SELECT " & vUsu.Codigo & ",codforpa,tipoformapago,descformapago " & Cad
            Cad = "INSERT INTO Usuarios.zsimulainm(codusu,codigo,conconam,nomconam) " & Cad
        
            Conn.Execute Cad
        End If
    End If
    
    
    
    'Voy a comprobar si ha metido algun registo
    espera 0.3
    SQL = "Select count(*) FROM  Usuarios.zpendientes where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not RS.EOF Then CONT = DBLet(RS.Fields(0), "N")
    RS.Close
    If CONT = 0 Then
        MsgBox "No se ha generado ningun valor(2)", vbExclamation
    Else
        CobrosPendientesCliente = True 'Para imprimir
    End If
    Exit Function
ECobrosPendientesCliente:
    MuestraError Err.Number, Err.Description
End Function



Private Function PagosPendienteProv(ByVal ListadoCuentas As String) As Boolean
'Dim MismaClavePrimaria As String


    On Error GoTo EPagosPendienteProv
    PagosPendienteProv = False
    
    'Trozo basico de PAGOS
    Cad = "FROM spagop ,cuentas ,sforpa,stipoformapago"
    Cad = Cad & " WHERE spagop.ctaprove = cuentas.codmacta"
    Cad = Cad & " AND spagop.codforpa = sforpa.codforpa"
    Cad = Cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    'Fecha
    RC = CampoABD(Text3(3), "F", "fecefect", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Text3(4), "F", "fecefect", False)
    If RC <> "" Then Cad = Cad & " AND " & RC

    'Cliente
    RC = CampoABD(txtCta(2), "T", "ctaprove", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtCta(3), "T", "ctaprove", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    'FORMA PAGO
    RC = CampoABD(txtFPago(6), "N", "spagop.codforpa", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtFPago(7), "N", "spagop.codforpa", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    
    
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        SQL = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then SQL = SQL & ","
                NumRegElim = 2
                SQL = SQL & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        Cad = Cad & " AND spagop.ctaprove IN (" & SQL & ")"
        
    End If
    
    
    'ORDEN
    Cad = Cad & " ORDER BY numfactu"
   
    
    
    
    'Contador
    SQL = "Select count(*) "
    SQL = SQL & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not RS.EOF Then
        'Total registros
        TotalRegistros = RS.Fields(0)
    End If
    RS.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    SQL = "SELECT spagop.*, cuentas.nommacta, stipoformapago.descformapago, stipoformapago.siglas,nomforpa " & Cad
    
    'Cambiamos
''''''    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''    'Compruebo si hay repetidos en fecfactu|numfactu|siglas
''''''    SQL = ""
''''''    MismaClavePrimaria = "|"
''''''    While Not RS.EOF
''''''        RC = Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu
''''''        If RC = SQL Then
''''''            RC = RC & "|"
''''''            If InStr(1, MismaClavePrimaria, "|" & RC) = 0 Then MismaClavePrimaria = MismaClavePrimaria & RC
''''''        Else
''''''            SQL = RC
''''''        End If
''''''        RS.MoveNext
''''''    Wend
''''''    SQL = RS.Source
''''''    RS.Close
    
    'Vuelvo a abrirlo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Agosto 2013
    'A�adimos en campo SITUACION donde pondra si esta emitido o no (emitdocum)
    
    'Mayo 2014
    'La factura la metemos en nomdirec. Asi NO da error duplicados
    
    CONT = 0
    SQL = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,nomdirec,"
    SQL = SQL & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,vencido,situacion) VALUES (" & vUsu.Codigo & ",'"
    Fecha = CDate(Text3(5).Text)
    DevfrmCCtas = ""
    While Not RS.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
'        'Por si se repiten misma fecfactura, mismo numero factura, mismo tipo de pago
'        If MismaClavePrimaria <> "" Then
'            'Hay claves repetidas no tiene en cuenta el vto
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "|"
'            'Enero 2011
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "#" & RS!numorden & "|"
'
'
'            If InStr(1, MismaClavePrimaria, RC) > 0 Then
'                RC = DevNombreSQL(RS!numfactu)
'                RC = FijaNumeroFacturaRepetido(RC)
'                Cad = RS!siglas & "','" & RC & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            Else
'                'El mismo de abajo
'                Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            End If
'        Else
'            'El procedimiento de antes
'            Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'        End If
'
        
        'mayo 2014
        Cad = RS!siglas & "','" & Format(CONT, "00000") & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden & ",'" & DevNombreSQL(RS!numfactu) & "'"
        
        
        'optMostraFP
        Cad = Cad & "," & RS!codforpa & ",'"
        If Me.optMostraFP(0).Value Then
            Cad = Cad & DevNombreSQL(RS!descformapago)
        Else
            Cad = Cad & DevNombreSQL(RS!nomforpa)
        End If
        Cad = Cad & "','" & RS!ctaprove & "','" & DevNombreSQL(RS!Nommacta) & "','"
        Cad = Cad & Format(RS!fecefect, FormatoFecha) & "',"
        Cad = Cad & TransformaComasPuntos(CStr(RS!ImpEfect)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(RS!imppagad) Then
            Cad = Cad & TransformaComasPuntos(CStr(RS!imppagad))
        Else
            Cad = Cad & "0"
        End If
        If Fecha > RS!fecefect Then
            Cad = Cad & ",1"
        Else
            Cad = Cad & ",0"
        End If
        
        'Agosto 2013
        'Si esta en un tal-pag
        Cad = Cad & ",'"
        If DBLet(RS!emitdocum, "N") > 0 Then Cad = Cad & "*"
        
        Cad = Cad & "')"  'lleva el apostrofe
        Cad = SQL & Cad
        Conn.Execute Cad
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
     
    PagosPendienteProv = True 'Para imprimir
    Exit Function
EPagosPendienteProv:
    MuestraError Err.Number, Err.Description
End Function



Private Function FijaNumeroFacturaRepetido(Numerofactura) As String
Dim I As Integer
Dim AUX As String
        If Len(Numerofactura) >= 10 Then
            MsgBox "Clave duplicada. Imposible insertar. " & RS!numfactu & ": " & RS!fecfactu, vbExclamation
            FijaNumeroFacturaRepetido = Numerofactura
            Exit Function
        End If
        
        'A�adiremos guienos por detras
        For I = Len(Numerofactura) To 10
            'A�adirenos espacios en blanco al final
            AUX = RS!numfactu & String(I - Len(Numerofactura), "_")
            If InStr(1, DevfrmCCtas, "|" & AUX & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = AUX
                If DevfrmCCtas = "" Then DevfrmCCtas = "|"
                DevfrmCCtas = DevfrmCCtas & AUX & "|"
                Exit Function
            End If
            
        Next I
        
        'Si llega aqui probaremos con los -
        For I = Len(Numerofactura) + 1 To 10
            'A�adirenos espacios en blanco al final
            AUX = String(I - Len(Numerofactura), "_") & RS!numfactu
            If InStr(1, DevfrmCCtas, "|" & AUX & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = AUX
                DevfrmCCtas = DevfrmCCtas & AUX & "|"
                Exit Function
            End If
            
        Next I
End Function


Private Sub txtNumero_GotFocus(Index As Integer)
    PonFoco txtNumero(Index)
End Sub



Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub


Private Sub txtnumfac_GotFocus(Index As Integer)
    PonFoco txtNumfac(Index)
End Sub

Private Sub txtnumfac_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtnumfac_LostFocus(Index As Integer)
    txtNumfac(Index).Text = Trim(txtNumfac(Index).Text)
    If txtNumfac(Index).Text = "" Then Exit Sub
    
    If Not IsNumeric(txtNumfac(Index).Text) Then
        MsgBox "Campo debe ser numerico.", vbExclamation
        txtNumfac(Index).Text = ""
        Ponerfoco txtNumfac(Index)
    End If
End Sub

Private Sub txtRem_GotFocus(Index As Integer)
    PonFoco txtRem(Index)
End Sub

Private Sub txtRem_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub


Private Sub txtSerie_GotFocus(Index As Integer)
    PonFoco txtSerie(Index)
End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
    txtSerie(Index).Text = UCase(txtSerie(Index))
End Sub

Private Sub txtVarios_GotFocus(Index As Integer)
    PonFoco txtVarios(Index)
End Sub

Private Sub txtVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub



Private Function ListadoRemesas() As Boolean
Dim AUX As String
    On Error GoTo EListadoRemesas
    ListadoRemesas = False
    
    SQL = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    Set RS = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /a�o desc                tipo   situacion
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    Cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe,haber) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del a�o 2005 ...
        CONT = RS!Anyo * 100000 + RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(RS!d1, "t") & "','" & DBLet(RS!descsituacion, "T") & "','"
        
        'Tipo remesa
        If RS!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf RS!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(RS!Importe)) & ",'" & Format(RS!fecremesa, FormatoFecha) & "')"
    
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
        If Me.chkRem(0).Value = 1 Then
            'fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,cuentaba
            RC = "SELECT fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,scobro.cuentaba,nommacta"
            RC = RC & " ,fecvenci,scobro.iban from scobro,cuentas where codrem=" & RS!Codigo & " AND anyorem =" & RS!Anyo
            RC = RC & " AND cuentas.codmacta = scobro.codmacta  ORDER BY "
            If Me.optOrdenRem(1).Value Then
                'Codmacta
                RC = RC & "scobro.codmacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(2).Value Then
                'Nommacta
                RC = RC & "nommacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(0).Value Then
                'Numero factura
                RC = RC & "numserie,codfaccl,fecfaccl"
            Else
                'fcto
                RC = RC & "fecvenci,numserie,codfaccl,fecfaccl"
            
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'If CONT = 200900004 Then Stop
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
                RC = RC & I & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                'Importe = miRsAux!impvenci - DBLet(miRsAux!impcobro, "N") + DBLet(miRsAux!Gastos, "N")
                If miRsAux!siturem > "B" Then
                    'No deberia ser NULL
                    Importe = DBLet(miRsAux!impcobro, "N")
                Else
                    Importe = miRsAux!impvenci + DBLet(miRsAux!Gastos, "N")
                End If
                RC = RC & Format(miRsAux!codfaccl, "000000000") & "','"
                
                'Aqui pondre el CCC para los efectos
                '---------------------------------------------------
                'rc=rc & codbanco,codsucur,digcontr,scobro.cuentaba
                AUX = ""
                If RS!Tiporem = 1 Then
                        If IsNull(miRsAux!codbanco) Then
                            AUX = "0000"
                        Else
                            AUX = Format(miRsAux!codbanco, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!codsucur) Then
                            AUX = AUX & "0000"
                        Else
                            AUX = AUX & Format(miRsAux!codsucur, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!digcontr) Then
                            AUX = AUX & "**"
                        Else
                            AUX = AUX & Format(miRsAux!digcontr, "00")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!cuentaba) Then
                            AUX = AUX & "0000"
                        Else
                            AUX = AUX & Format(miRsAux!cuentaba, "0000000000")
                        End If
                Else
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                 
                End If
                
                'Nuevo ENERO 2010
                'Fecha vto
                AUX = DBLet(miRsAux!IBAN, "T") & AUX
                If Len(AUX) > 24 Then AUX = Mid(AUX, 1, 24)
                AUX = AUX & "|" & Format(miRsAux!fecvenci, "dd/mm/yy")
                
                RC = RC & AUX & "'," & TransformaComasPuntos(CStr(Importe))
                
                'En el haber pongo el ascii de la serie
                '--------------------------------------
                RC = RC & "," & Asc(Left(DBLet(miRsAux!NUmSerie, "T") & " ", 1)) & ")"
                RC = Cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
        Else
            'Tenemos k insertar una unica linea a blancos
            RC = vUsu.Codigo & "," & CONT & ",'1999-12-31'," & I & ",'','','','',0,0)"
            RC = Cad & RC
            
            Conn.Execute RC
        End If
        TotalRegistros = TotalRegistros + 1
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesas = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing

End Function









Private Function ListadoRemesasBanco() As Boolean
Dim AUX As String
Dim Cad2 As String
Dim J As Integer
    On Error GoTo EListadoRemesas
    ListadoRemesasBanco = False
    
    SQL = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    Set RS = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /a�o desc                tipo   situacion
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1,observa1) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del a�o 2005 ...
        CONT = RS!Anyo * 100000 + RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(RS!d1, "t") & "','" & DBLet(RS!descsituacion, "T") & "','"
        
        'Tipo remesa
        If RS!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf RS!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(RS!Importe)) & ",'" & Format(RS!fecremesa, FormatoFecha) & "','"
        
        Cad2 = "Select * from ctabancaria where codmacta = '" & RS!codmacta & "'"
        miRsAux.Open Cad2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad2 = "NO ENCONTRADO"
        If Not miRsAux.EOF Then
            Cad2 = Trim(DBLet(miRsAux!IBAN, "T") & " ") & Format(DBLet(miRsAux!Entidad, "N"), "0000") & " " & Format(DBLet(miRsAux!oficina, "N"), "0000") & " "
            If IsNull(miRsAux!Control) Then
                Cad2 = Cad2 & "*"
            Else
                Cad2 = Cad2 & miRsAux!Control
            End If
            Cad2 = Cad2 & " " & Format(DBLet(miRsAux!CtaBanco, "N"), "0000000000")
        End If
        miRsAux.Close
        RC = RC & Cad2 & "')"
        'ctabancaria(entidad,oficina,control,ctabanco)
        Cad2 = ""
        
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
        
            'Voy a comprobar que existen
            RC = "SELECT codmacta,reftalonpag FROM scobro "
            RC = RC & "  WHERE codrem=" & RS!Codigo & " AND anyorem =" & RS!Anyo
            RC = RC & " GROUP BY 1,2 "
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad2 = ""
            While Not miRsAux.EOF
                Cad2 = Cad2 & " scarecepdoc.codmacta = '" & miRsAux!codmacta & "' AND numeroref = '" & DevNombreSQL(miRsAux!reftalonpag) & "'|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If Cad2 = "" Then
                MsgBox "Error obteniendo cuenta/referenciatalon", vbExclamation
                RS.Close
                Exit Function
            End If
                
            'Comprobare que existen y de paso los inserto
            While Cad2 <> ""
                J = InStr(1, Cad2, "|")
                AUX = Mid(Cad2, 1, J - 1)
                Cad2 = Mid(Cad2, J + 1)
                
                'RC = "SELECT * FROM scarecepdoc ,slirecepdoc,cuentas WHERE codigo=id AND scarecepdoc.codmacta=cuentas.codmacta AND " & Aux
                RC = "SELECT * FROM scarecepdoc ,cuentas WHERE  scarecepdoc.codmacta=cuentas.codmacta AND " & AUX
               
                miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "No se encuentra la referencia; " & AUX, vbExclamation
                    miRsAux.Close
                    RS.Close
                    Exit Function
                End If
                
                While Not miRsAux.EOF
            
                
                
                
                
                    
                    'Para insertar en la otra
                    Cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu,  nommacta,codmacta, numdocum, ampconce, debe,haber) VALUES ("
                
                    'Trampas:  Entre codmacta numdocum llevare el numero de talon. Ya que suman 20 y reftal es len20
                    RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fechavto, FormatoFecha) & "',"
                    RC = RC & I & ",'" & DevNombreSQL(miRsAux!Nommacta) & "','"
                    Importe = DBLet(miRsAux!Importe, "N")
                    
                    'Referencia talon
                    AUX = DevNombreSQL(miRsAux!numeroref)
                    If Len(AUX) > 10 Then
                        RC = RC & Mid(AUX, 1, 10) & "','" & Mid(AUX, 11)
                    Else
                        RC = RC & AUX & "','"
                    End If
                    'Banco
                    RC = RC & "','" & DevNombreSQL(miRsAux!banco) & "',"
                    
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                    RC = RC & TransformaComasPuntos(CStr(Importe))
                    
                    'En el haber pongo el ascii de la serie
                    '--------------------------------------
                    RC = RC & ",0)"
                    RC = Cad & RC
                
                    Conn.Execute RC
                
                    'Sig
                    I = I + 1
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
            Wend

        TotalRegistros = TotalRegistros + 1
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    
    
      Set RS = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesasBanco = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing

End Function




'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'               CREDITO CAUCION
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Private Function ListadoTransferencias() As Boolean
Dim Importe As Currency

    On Error GoTo EListadoTransferencias
    ListadoTransferencias = False
    
    SQL = ""
    RC = CampoABD(txtNumero(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtNumero(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    
    Cad = RC
    
    Set RS = New ADODB.Recordset
    
    RC = "SELECT stransfer.*,nommacta from stransfer"
    If Opcion = 13 Then RC = RC & "cob"
    RC = RC & " as stransfer,cuentas "
    RC = RC & " WHERE stransfer.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna valor para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    If Opcion = 13 Then Conn.Execute "Delete from usuarios.zcuentas where codusu =" & vUsu.Codigo
        
    
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /a�o desc
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    Cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe) VALUES ("
    
    
    

    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del a�o 2005 ...
        CONT = RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        RC = RC & TransformaComasPuntos("0") & ",'" & Format(RS!Fecha, FormatoFecha) & "')"
    
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
     
            
            If Opcion = 13 Then
                RC = "scobro"
            Else
                RC = "spagop"
            End If
            RC = "SELECT " & RC & ".*,nommacta from cuentas," & RC
            RC = RC & " WHERE transfer = " & RS!Codigo
            RC = RC & " AND cuentas.codmacta = "
            If Opcion = 13 Then
                RC = RC & " scobro.codmacta "
                RC = RC & " ORDER BY scobro.codmacta,fecfaccl"
            Else
                RC = RC & " spagop.ctaprove "
                RC = RC & " ORDER BY ctaprove,fecfactu"
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe "
                If Opcion = 13 Then
                    Fecha = miRsAux!fecfaccl
                Else
                    Fecha = miRsAux!fecfactu
                End If
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(Fecha, FormatoFecha) & "',"
                RC = RC & I & ",'"
                If Opcion = 13 Then
                    RC = RC & miRsAux!codmacta
                Else
                    RC = RC & miRsAux!ctaprove
                End If
                
                RC = RC & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                
                
                'Cuenta
                If Opcion <> 13 Then
                    RC = RC & DevNombreSQL(miRsAux!numfactu) & "','"
                    
                    'Noviembre 2013
                    'A�adimos el IBAN
                    
                    RC = RC & Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!Entidad, "T"), "0000")) & " " & Format(DBLet(miRsAux!oficina, "T"), "0000") & " "
                    RC = RC & Mid(DBLet(miRsAux!CC, "T") & "**", 1, 2) & " " & Right(String(10, "0") & DBLet(miRsAux!cuentaba, "T"), 10)
                    Importe = miRsAux!ImpEfect - (DBLet(miRsAux!imppagad, "N"))
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                Else
                    RC = RC & DevNombreSQL(miRsAux!codfaccl) & "','"
                    
                    CadenaDesdeOtroForm = "NO"
                    If DBLet(miRsAux!codbanco, "N") > 0 Then
                        'Es especifico para ESCALONO, pero no molesta a nadie
                        If DBLet(miRsAux!cuentaba, "T") = "8888888888" Then
                            'Seguira poniendo  cuenta no domiciliada
                        Else
                            CadenaDesdeOtroForm = ""
                        End If
                    End If
                    If CadenaDesdeOtroForm = "" Then
                        'OK, ponemos la cuenta
                        CadenaDesdeOtroForm = Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!codbanco, "N"), "0000")) & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "**  ******" & Right(String(4, "0") & DBLet(miRsAux!cuentaba, "T"), 4)
                        
                    Else
                        'CUENTANODOMICILIADA
                        CadenaDesdeOtroForm = "NODOM"  'en el rpt haremos un replace
                    End If
                    RC = RC & CadenaDesdeOtroForm
                    Importe = miRsAux!impvenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                End If
                RC = Cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
'        Else
'            'Tenemos k insertar una unica linea a blancos
'            RC = vUsu.Codigo & "," & CONT & ",''," & I & ",'','','','',0)"
'            RC = Cad & RC
'
'            Conn.Execute RC
'        End If
        RS.MoveNext
    Wend
    RS.Close
    CadenaDesdeOtroForm = ""
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If Opcion = 13 Then
        'Puede ser carta
        If chkCartaAbonos.Value Then
            'En nommacta pongo la provincia  (desprovi)
            Cad = "INSERT INTO usuarios.zcuentas(codusu,codmacta,nommacta,razosoci,dirdatos,codposta,despobla,nifdatos)"
            Cad = Cad & " Select " & vUsu.Codigo & ",codmacta,desprovi,razosoci,dirdatos,codposta,despobla,nifdatos FROM cuentas WHERE "
            Cad = Cad & " codmacta IN (select distinct(codmacta) from usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo & ")"
            Ejecuta Cad
        
        
            Cad = "apoderado"
            RC = DevuelveDesdeBD("contacto", "empresa2", "1", "1", "N", Cad)
            If RC = "" Then RC = Cad
            If RC <> "" Then
                Cad = "UPDATE usuarios.ztesoreriacomun SET observa1='" & DevNombreSQL(RC) & "'"
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo
                Conn.Execute Cad
            End If
        End If
    End If
    
    If Me.chkRem(0).Value = 1 Then
        If I = 1 Then
            MsgBox "No hay vencimientos asociados a las transferencias", vbExclamation
            Exit Function
        End If
    End If
    ListadoTransferencias = True
    Exit Function
EListadoTransferencias:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing
End Function





Private Function ListAseguBasico() As Boolean
    On Error GoTo EListAseguBasico
    ListAseguBasico = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "Select * from cuentas where numpoliz<>"""""
    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecsolic", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecconce", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY " & RC
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,"
    Cad = Cad & "importe2,observa1,observa2) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DBLet(miRsAux!nifdatos, "T") & "','"
        SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Fecha sol y concesion
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecsolic, "F", True) & "," & CampoBD_A_SQL(miRsAux!fecconce, "F", True) & ","
        'Importes sol y concesion
        SQL = SQL & CampoBD_A_SQL(miRsAux!credisol, "N", True) & "," & CampoBD_A_SQL(miRsAux!credicon, "N", True) & ","
        'Observaciones
        RC = Memo_Leer(miRsAux!observa)
        If Len(RC) = 0 Then
            'Los dos campos NULL
            SQL = SQL & "NULL,NULL"
        Else
            If Len(RC) < 255 Then
                SQL = SQL & "'" & DevNombreSQL(RC) & "',NULL"
            Else
                SQL = SQL & "'" & DevNombreSQL(Mid(RC, 1, 255))
                RC = Mid(RC, 256)
                SQL = SQL & "','" & DevNombreSQL(Mid(RC, 1, 255)) & "'"
            End If
        End If
        
        SQL = SQL & ")"
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAseguBasico = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAseguBasico:
    MuestraError Err.Number, "ListAseguBasico"
End Function





Private Function ListAsegFacturacion() As Boolean
Dim FP As Integer 'Forma de pago
Dim Cadpago As String
    On Error GoTo EListAsegFacturacion
    ListAsegFacturacion = False
    
    Cad = "DELETE FROM Usuarios.zpendientes  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    
    If Me.optFecgaASig(0).Value Then
        Cad = "fecfaccl"
    Else
        Cad = "fecvenci"
    End If
        
    SQL = ""
    RC = CampoABD(Text3(21), "F", Cad, True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", Cad, False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    
    
    Cad = "Select scobro.*,nommacta,numpoliz,nomforpa,forpa from scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY " & RC
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0

    Cad = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    Cad = Cad & "codforpa, nomforpa, codmacta, nombre, fecVto, importe,"
    Cad = Cad & "Situacion,pag_cob, vencido,  gastos) VALUES (" & vUsu.Codigo & ","
    Cadpago = ","
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = "'" & miRsAux!NUmSerie & "','" & Format(miRsAux!codfaccl, "000000000") & "','" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
        FP = miRsAux!codforpa
        If optFP(1).Value Then
            If DBLet(miRsAux!ForPa, "N") > 0 Then
                FP = miRsAux!ForPa
                If InStr(1, Cadpago, "," & FP & ",") = 0 Then Cadpago = Cadpago & FP & ","
            End If
        End If
        SQL = SQL & miRsAux!numorden & "," & FP & ",'" & DevNombreSQL(miRsAux!nomforpa) & "','" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta)
        SQL = SQL & "','" & Format(miRsAux!fecvenci, FormatoFecha) & "',"
        'IMporte
        Importe = miRsAux!impvenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        'Situacion tengo numpoliza
        SQL = SQL & ",'" & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Gastos e imvenci van a la columna pag_cob   Julio 2009
        Importe = miRsAux!impvenci + DBLet(miRsAux!Gastos, "N")
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        'El resto
        SQL = SQL & ",0,NULL)"
        
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If CONT = 0 Then
        MsgBox "Ningun datos con esos valores", vbExclamation
        Exit Function
    End If
    
    
    'Si ha cambiado valores en la forma de pago
    If Len(Cadpago) > 1 Then
        Cadpago = Mid(Cadpago, 2)
        Cadpago = Mid(Cadpago, 1, Len(Cadpago) - 1)
        Cad = "select codforpa,nomforpa from sforpa where codforpa in (" & Cadpago & ") GROUP by  codforpa"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = " WHERE codusu = " & vUsu.Codigo & " AND codforpa = "
        While Not miRsAux.EOF
            SQL = "UPDATE Usuarios.zpendientes SET nomforpa = '" & DevNombreSQL(miRsAux!nomforpa) & "'" & Cad & miRsAux!codforpa
            If Not Ejecuta(SQL) Then MsgBox "Error actualizando tmp.  Forpa: " & miRsAux!codforpa & " " & miRsAux!nomforpa, vbExclamation
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    ListAsegFacturacion = True
    
    
    Exit Function
EListAsegFacturacion:
    MuestraError Err.Number, "ListAseguBasico"
End Function



Private Function ListAsegImpagos() As Boolean
    On Error GoTo EListAsegImpagos
    ListAsegImpagos = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,scobro.codmacta,nommacta,despobla,desprovi,numpoliz,nomforpa from "
    Cad = Cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    'Impagados
    Cad = Cad & " AND devuelto = 1"
    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY " & RC
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,texto5,texto6,fecha1,importe1) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DevNombreSQL(DBLet(miRsAux!despobla, "T")) & "','"
        SQL = SQL & DevNombreSQL(DBLet(miRsAux!desprovi, "T")) & "','" & DevNombreSQL(miRsAux!numpoliz) & "','"
        SQL = SQL & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecvenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!impvenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        
    
        SQL = SQL & ")"
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegImpagos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegImpagos:
    MuestraError Err.Number, "ListAsegImpagos"
End Function


Private Function ListAsegEfectos() As Boolean
Dim TotalCred As Currency

    On Error GoTo EListAsegEfectos
    ListAsegEfectos = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,credicon from "
    Cad = Cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba

    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY codmacta,fecvenci"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,fvto,impvto,disponible,vencida
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,importe2,opcion) VALUES (" & vUsu.Codigo & ","
    RC = ""
    
    While Not miRsAux.EOF
        If RC <> miRsAux!codmacta Then
            RC = miRsAux!codmacta
            TotalCred = DBLet(miRsAux!credicon, "N")
            CadenaDesdeOtroForm = ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            If IsNull(miRsAux!credicon) Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0,00','"
            Else
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(miRsAux!credicon, FormatoImporte) & "','"
            End If
        End If
        CONT = CONT + 1
        SQL = CONT & CadenaDesdeOtroForm
        SQL = SQL & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha fac
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecvenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!impvenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        TotalCred = TotalCred - Importe
        SQL = SQL & "," & TransformaComasPuntos(CStr(TotalCred))
       
        'Devuelto
        SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegEfectos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "ListAsegEfec"
End Function



Private Sub GeneraComboCuentas()
'
'    If Opcion = 1 Then
'        'COBROS PENDIENTES
'    Else: Pagos
'
        cmbCuentas(Opcion - 1).Clear
        cmbCuentas(Opcion - 1).AddItem "Sin especificar"
        
        cmbCuentas(Opcion - 1).AddItem "Crear selecci�n"
              
        'En el tag tendremos las cuentas seleccionadas
        If Me.cmbCuentas(Opcion - 1).Tag <> "" Then cmbCuentas(Opcion - 1).AddItem "Cuentas seleccionadas"


    'Cargo aqui tb los checks
    CargaTextosTipoPagos False
End Sub



Private Sub CargaTextosTipoPagos(Reclamaciones As Boolean)
    
    Set miRsAux = New ADODB.Recordset
    Cad = "Select tipoformapago, descformapago,siglas from stipoformapago order by tipoformapago "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Reclamaciones Then
            chkTipPagoRec(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPagoRec(miRsAux!tipoformapago).Visible = True
            chkTipPagoRec(miRsAux!tipoformapago).Value = 1
        
        Else
            chkTipPago(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPago(miRsAux!tipoformapago).Visible = True
            chkTipPago(miRsAux!tipoformapago).Value = 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)


    Select Case OpcionGrid
    Case 0
    'Diario
        DevfrmCCtas = "0"
        Cad = "N�mero|numdiari|N|30�"
        Cad = Cad & "Descripci�n|desdiari|T|60�"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "Tiposdiario"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Diario"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
           Me.txtDiario(Indice) = RecuperaValor(DevfrmCCtas, 1)
           Me.txtDescDiario(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
 Case 1
        'Conceptos
        DevfrmCCtas = "0"
        Cad = "Codigo|codconce|N|30�"
        Cad = Cad & "Descripci�n|nomconce|T|60�"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "Conceptos"
        frmB.vSQL = ""
        
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "CONCEPTOS"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
           Me.txtConcpto(Indice) = RecuperaValor(DevfrmCCtas, 1)
           Me.txtDescConcepto(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
        
    Case 2
        'Centros de coste
        DevfrmCCtas = "0"
        Cad = "Codigo|codccost|T|30�"
        Cad = Cad & "Descripci�n|nomccost|T|60�"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "cabccost"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Centros de coste"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
            
           txtCCost(Indice) = RecuperaValor(DevfrmCCtas, 1)
           txtDescCCoste(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
        
    Case 3
        'Gasto fijos  sgastfij codigo Descripcion
        DevfrmCCtas = "0"
        Cad = "C�digo|codigo|T|30�"
        Cad = Cad & "Descripci�n|Descripcion|T|60�"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "sgastfij"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Gastos fijos"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
            
           txtGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 1)
           txtDescGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
        
    Case 4
        'Gasto fijos  sgastfij codigo Descripcion
        '-------------------------------------------
        DevfrmCCtas = "0"
        Cad = "C�digo|codigo|T|10�"
        Cad = Cad & "Fecha|fecha|T|26�"
        Cad = Cad & "Cuenta|codmacta|T|14�"
        Cad = Cad & "Nombre|nommacta|T|50�"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "scompenclicab"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = "Hco. compensaciones cliente"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
            DevfrmCCtas = RecuperaValor(DevfrmCCtas, 1)
            If DevfrmCCtas = "" Then DevfrmCCtas = "0"
            If Val(DevfrmCCtas) Then
                CONT = Val(DevfrmCCtas)
                ImprimiCompensacion CONT
            End If
           
        End If
    End Select
End Sub

                                       '                Para saber el index del listview
Public Sub InsertaItemComboCompensaVto(TEXTO As String, Indice As Integer)
    Me.cboCompensaVto.AddItem TEXTO
    Me.cboCompensaVto.ItemData(Me.cboCompensaVto.NewIndex) = Indice
End Sub


Private Function GeneraDatosTalPag() As Boolean
Dim B As Boolean

    'Borramos los tmp
    SQL = "DELETE FROM usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    If chkLstTalPag(3).Value = 1 Then
        B = GeneraDatosTalPagDesglosado
    Else
        'Sin desglosar
        B = GeneraDatosTalPagSinDesglose
    End If
    GeneraDatosTalPag = B
End Function

Private Function GeneraDatosTalPagDesglosado() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagDesglosado = False
    
    

       
       
    SQL = "select slirecepdoc.*,scarecepdoc.*,nommacta,nifdatos from slirecepdoc,scarecepdoc,cuentas "
    SQL = SQL & " where slirecepdoc.id =scarecepdoc.codigo and scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    If cboListPagare.ListIndex >= 1 Then SQL = SQL & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    '------------------------------------------------------------------------
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then SQL = SQL & " AND talon = " & I

    'Si ID
    If txtNumfac(2).Text <> "" Then SQL = SQL & " AND codigo >= " & txtNumfac(2).Text
    If txtNumfac(3).Text <> "" Then SQL = SQL & " AND codigo <= " & txtNumfac(3).Text

    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not RS.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        SQL = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        SQL = SQL & "'" & DevNombreSQL(RS!numeroref) & "','" & DevNombreSQL(RS!banco) & "','"
        SQL = SQL & DevNombreSQL(RS!codmacta) & "','" & DevNombreSQL(RS!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        SQL = SQL & RS!NUmSerie & Format(RS!numfaccl, "000000") & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs!Importe)) & ","
        SQL = SQL & TransformaComasPuntos(CStr(RS.Fields(5))) & ",'"   'La columna 5 es sli.importe
        
        'texto6=nifdatos
        SQL = SQL & DevNombreSQL(DBLet(RS!nifdatos, "N"))
        
        '`fecha1`,`fecha2`,`fecha3`
        SQL = SQL & "','" & Format(RS!fecharec, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fechavto, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fecfaccl, FormatoFecha) & "')"
    
        RC = RC & SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        SQL = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        SQL = SQL & "`texto4`,`texto5`,`importe1`,texto6,`fecha1`,`fecha2`,`fecha3`) VALUES "
        SQL = SQL & RC
        Conn.Execute SQL
    
        'Si estamos emitiendo el justicante de recepcion, guardare en z340 los campos
        'fiscales del cliente para su impresion
        If Me.chkLstTalPag(2).Value = 1 Then
            SQL = "DELETE FROM usuarios.z347 WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            SQL = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            espera 0.3
            
            
            'En texto3 esta la codmacta
            SQL = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY texto3"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = ""
            While Not RS.EOF
                RC = RC & ", '" & RS!texto3 & "'"
                RS.MoveNext
            Wend
            RS.Close
            
            
            
            
            
            'No puede ser EOF
            RC = Trim(Mid(RC, 2))
            'Monto un superselect
            'pongo el IGNORE por si acaso hay cuentas con el mismo NIF
            SQL = "insert ignore into usuarios.z347 (`codusu`,`cliprov`,`nif`,`razosoci`,`dirdatos`,`codposta`,`despobla`,`Provincia`)"
            SQL = SQL & " SELECT " & vUsu.Codigo & ",0,nifdatos,razosoci,dirdatos,codposta,despobla,desprovi FROM cuentas where codmacta in (" & RC & ")"
            Conn.Execute SQL
    
    
    
            'Ahora meto los datos de la empresa
            Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir,"
            Cad = Cad & "contacto) VALUES ("
            Cad = Cad & vUsu.Codigo
                
                
            'Monta Datos Empresa
            RS.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
            If RS.EOF Then
                MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
                RC = ",'','','','','',''"  '6 campos
            Else
                RC = DBLet(RS!siglasvia) & " " & DBLet(RS!direccion) & "  " & DBLet(RS!numero) & ", " & DBLet(RS!puerta)
                RC = ",'" & DBLet(RS!nifempre) & "','" & vEmpresa.nomempre & "','" & RC & "','"
                RC = RC & DBLet(RS!codpos) & "','" & DBLet(RS!Poblacion) & "','" & DBLet(RS!provincia) & "','" & DBLet(RS!contacto) & "')"
            End If
            RS.Close
            Cad = Cad & RC
            Conn.Execute Cad
            
            
            
    
        End If
        GeneraDatosTalPagDesglosado = True
    Else
        'I>0
        MsgBox "No hay datos", vbExclamation
    End If

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function



Private Function GeneraDatosTalPagSinDesglose() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagSinDesglose = False
    
    

       
       
    SQL = "select scarecepdoc.*,nommacta from scarecepdoc,cuentas "
    SQL = SQL & " where  scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    'SQL = SQL & " AND LlevadoBanco = " & Abs(chkLstTalPag(2).Value)
    If cboListPagare.ListIndex >= 1 Then SQL = SQL & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then SQL = SQL & " AND talon = " & I
    'Si ID
    If txtNumfac(2).Text <> "" Then SQL = SQL & " AND codigo >= " & txtNumfac(2).Text
    If txtNumfac(3).Text <> "" Then SQL = SQL & " AND codigo <= " & txtNumfac(3).Text



    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not RS.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        SQL = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        SQL = SQL & "'" & DevNombreSQL(RS!numeroref) & "','" & DevNombreSQL(RS!banco) & "','"
        SQL = SQL & DevNombreSQL(RS!codmacta) & "','" & DevNombreSQL(RS!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        SQL = SQL & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs.Fields(5))) & ","   'La columna 5 es sli.importe
        SQL = SQL & TransformaComasPuntos(CStr(RS!Importe)) & ","
        
        '
        '`fecha1`,`fecha2`,`fecha3`
        SQL = SQL & "'" & Format(RS!fecharec, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fechavto, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(Now, FormatoFecha) & "')"
    
        RC = RC & SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        SQL = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        SQL = SQL & "`texto4`,`texto5`,`importe1`,`fecha1`,`fecha2`,`fecha3`) VALUES "
        SQL = SQL & RC
        Conn.Execute SQL
        GeneraDatosTalPagSinDesglose = True
    Else
        MsgBox "No hay datos", vbExclamation
    End If
    
    

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function





Private Function ListadoOrdenPago() As Boolean
    On Error GoTo EListadoOrdenPago
    ListadoOrdenPago = False

    'Borramos
    Cad = "DELETE from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad
    Set miRsAux = New ADODB.Recordset
    'Inserttamos
    RC = ""
    If txtCtaBanc(3).Text <> "" Or txtCtaBanc(4).Text <> "" Then
        If txtCtaBanc(3).Text <> "" Then RC = " codmacta >= '" & txtCtaBanc(3).Text & "'"
        
        If txtCtaBanc(4).Text <> "" Then
            If RC <> "" Then RC = RC & " AND "
            RC = RC & " codmacta <= '" & txtCtaBanc(4).Text & "'"
        End If
        Cad = "Select codmacta from ctabancaria where " & RC
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        RC = ""
        While Not miRsAux.EOF
            RC = RC & ", '" & miRsAux!codmacta & "'"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If RC = "" Then
            MsgBox "Ning�n banco con esos valores", vbExclamation
            Exit Function
        End If
           
        RC = Mid(RC, 2)
    End If
    
    
    SQL = ""
    If Text3(26).Text <> "" Then SQL = SQL & " AND fecefect >= '" & Format(Text3(26).Text, FormatoFecha) & "'"
    If Text3(27).Text <> "" Then SQL = SQL & " AND fecefect <= '" & Format(Text3(27).Text, FormatoFecha) & "'"
    If RC <> "" Then SQL = SQL & " AND ctabanc1 in (" & RC & ")"
    
    
    'JULIO2013
    'La variable contdocu grabaremos emitdocum, y en el listado sabremos si es de talon/pagere
    'para poder separalos
    'Antes: contdocu, ahora emitdocum
    
    'Agosto 2014
    'Tipo de pago
    Cad = "select " & vUsu.Codigo & ",`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,`impefect`-coalesce(imppagad,0),"
    Cad = Cad & " `ctabanc1`,`ctabanc2`,`emitdocum`,spagop.entidad,spagop.oficina,spagop.CC,spagop.cuentaba,"
    Cad = Cad & " nommacta,'error','error',descformapago "
    
    Cad = Cad & " from spagop,cuentas,sforpa,stipoformapago "
    Cad = Cad & " WHERE spagop.ctaprove = cuentas.codmacta AND spagop.codforpa=sforpa.codforpa and tipoformapago=tipforpa"
    'Ponemos un check si salen negativos o no
    RC = " AND impefect >=0"
    If Me.chkPagBanco(0).Value = 1 And Me.chkPagBanco(1).Value = 1 Then RC = "" 'Salen todos
    Cad = Cad & RC 'todos o solo positivos
    Cad = Cad & SQL
    
    SQL = "INSERT INTO usuarios.zlistadopagos (`codusu`,`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,"
    SQL = SQL & " `impefect`,`ctabanc1`,`ctabanc2`,`contdocu`,`entidad`,`oficina`,`CC`,`cuentaba`,"
    SQL = SQL & " `nomprove`,`nombanco`,`cuentabanco`,TipoForpa) " & Cad
    Conn.Execute SQL
    
    Cad = DevuelveDesdeBD("count(*)", "usuarios.zlistadopagos", "codusu", vUsu.Codigo)
    If Val(Cad) = 0 Then
        MsgBox "Ningun vencimiento con esos valores", vbExclamation
        Exit Function
    End If
    
    'Actualizo los datos de los bancos `nombanco`,`cuentabanco`
    Cad = "Select ctabanc1 from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    Cad = Cad & " AND ctabanc1 <>'' GROUP BY ctabanc1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & miRsAux!ctabanc1 & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    While Cad <> ""
        I = InStr(1, Cad, "|")
        If I = 0 Then
            Cad = ""
        Else
            RC = Mid(Cad, 1, I - 1)
            Cad = Mid(Cad, I + 1)
            
            SQL = "Select ctabancaria.codmacta,ctabancaria.entidad, ctabancaria.oficina, ctabancaria.control, ctabancaria.ctabanco,"
            SQL = SQL & " ctabancaria.descripcion,nommacta from  ctabancaria,cuentas where ctabancaria.codmacta=cuentas.codmacta "
            SQL = SQL & " AND ctabancaria.codmacta ='" & RC & "'"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                SQL = "Cuenta banco erronea: " & vbCrLf & "Hay vencimientos asociados a la cuenta " & RC & " que no esta en bancos"
                MsgBox SQL, vbExclamation
            Else
                SQL = DBLet(miRsAux!Descripcion, "T")
                If SQL = "" Then SQL = miRsAux!Nommacta
                SQL = DevNombreSQL(SQL) & "|"
                
                'Enti8dad...
                I = DBLet(miRsAux!Entidad, "0")
                SQL = SQL & Format(I, "0000")
                                'Oficina...
                I = DBLet(miRsAux!oficina, "0")
                SQL = SQL & Format(I, "0000")
                                'CC...
                RC = DBLet(miRsAux!Control, "T")
                If RC = "" Then RC = "**"
                SQL = SQL & RC
                'cuenta
                RC = DBLet(miRsAux!CtaBanco, "T")
                If RC = "" Then RC = "    **"
                SQL = SQL & RC & "|"
                
                
                RC = "UPDATE usuarios.zlistadopagos set `nombanco`='" & RecuperaValor(SQL, 1)
                RC = RC & "',`cuentabanco`='" & RecuperaValor(SQL, 2) & "'"
                RC = RC & " WHERE ctabanc1 = '" & miRsAux!codmacta & "' AND codusu = " & vUsu.Codigo
                Conn.Execute RC
                
            End If
            miRsAux.Close
        End If
    Wend
    
    ListadoOrdenPago = True
    Set miRsAux = Nothing
    Exit Function
EListadoOrdenPago:
    MuestraError Err.Number, "ListadoOrdenPago"
End Function



Private Function ListadoReclamas() As Boolean

On Error GoTo EListadoReclamas

    ListadoReclamas = False
        

    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = ""
    Cad = ""
    
    If Text3(28).Text <> "" Or Text3(29).Text <> "" Then
        RC = DesdeHasta("F", 28, 29, "F.Reclama")
        If RC <> "" Then Cad = Cad & " " & RC
            
        RC = CampoABD(Text3(28), "F", "fecreclama", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(Text3(29), "F", "fecreclama", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    
    
    If txtCta(15).Text <> "" Or txtCta(16).Text <> "" Then
        RC = DesdeHasta("C", 15, 16, "Cta")
        If RC <> "" Then Cad = Cad & " " & RC
            
        RC = CampoABD(txtCta(15), "T", "codmacta", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(txtCta(16), "T", "codmacta", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    SQL = "Select * from shcocob" & SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`texto4`,`texto5`,`texto6`,`importe1`,`importe2`,`fecha1`,`fecha2`,"
    RC = RC & "`fecha3`,`texto`,`observa2`,`opcion`) VALUES "
    SQL = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = SQL & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & RS!codmacta & "','"
        'text 2 y 3
        SQL = SQL & DevNombreSQL(RS!Nommacta) & "','" & RS!NUmSerie & Format(RS!codfaccl, "000000") & "','"
        '4 y 5
        SQL = SQL & RS!numorden & "','"
        If Val(RS!carta) = 1 Then
            SQL = SQL & "Email"
        ElseIf Val(RS!carta) = 2 Then
            SQL = SQL & "Tel�fono"
        Else
            SQL = SQL & "Carta"
        End If
        'Text6, importe 1 y 2
        SQL = SQL & "',''," & TransformaComasPuntos(CStr(RS!impvenci)) & ",NULL,"
        'Fec1 reclama fec2 factra   fec3
        SQL = SQL & "'" & Format(RS!fecreclama, FormatoFecha) & "','" & Format(RS!fecfaccl, FormatoFecha) & "',NULL,"
        DevfrmCCtas = Memo_Leer(RS!observaciones)
        If DevfrmCCtas = "" Then
            DevfrmCCtas = "NULL"
        Else
            DevfrmCCtas = "'" & DevNombreSQL(DevfrmCCtas) & "'"
        End If
        SQL = SQL & DevfrmCCtas & ",NULL,0)"


        'Siguiente
        RS.MoveNext
        
        
        If Len(SQL) > 100000 Then
            SQL = Mid(SQL, 2) 'QUITO LA COMA
            SQL = RC & SQL
            Conn.Execute SQL
            SQL = ""
        End If
            
        
    Wend
    RS.Close
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'QUITO LA COMA
            SQL = RC & SQL
            Conn.Execute SQL
        End If
        
        
    If NumRegElim > 0 Then
        ListadoReclamas = True
    Else
        MsgBox "Ningun dato devuelto", vbExclamation
    End If
    Exit Function
EListadoReclamas:
    MuestraError Err.Number, Err.Description
End Function





'******************************************************************************************
'
'   Listado gastos fijos

Private Function ListadoGastosFijos() As Boolean

On Error GoTo EListadoGF

    ListadoGastosFijos = False
        

    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = ""
    Cad = ""
    
    
    DevfrmCCtas = "" ' ON del left join , NO al WHERE
    If Text3(30).Text <> "" Or Text3(31).Text <> "" Then
        RC = DesdeHasta("F", 30, 31, "Fecha")
        If RC <> "" Then Cad = Cad & " " & Trim(RC)
            
        RC = CampoABD(Text3(30), "F", "fecha", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(Text3(31), "F", "fecha", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    DevfrmCCtas = SQL
    SQL = ""
    
    'Este si que va al where
    If txtGastoFijo(0).Text <> "" Or txtGastoFijo(1).Text <> "" Then
        RC = DesdeHasta("GF", 0, 1, "Codigo")
        If RC <> "" Then
            If Cad <> "" Then
                'Ya esta la fecha
                If Len(Cad & RC) > 100 Then Cad = Cad & """ + chr(13) + """
            End If
            Cad = Cad & " " & Trim(RC)
        End If
            
        RC = CampoABD(txtGastoFijo(0), "N", "sgastfij.codigo", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(txtGastoFijo(1), "N", "sgastfij.codigo", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    
   
   
    RC = " FROM sgastfij left join sgastfijd ON sgastfij.Codigo = sgastfijd.Codigo"
    If DevfrmCCtas <> "" Then RC = RC & " AND " & DevfrmCCtas
    If SQL <> "" Then RC = RC & " WHERE " & SQL
    SQL = "SELECT sgastfij.codigo,descripcion,ctaprevista,fecha,importe" & RC
    
    

    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`importe1`,`fecha1`) VALUES "
    SQL = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = SQL & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & Format(RS!Codigo, "00000") & "','"
        'text 2 y 3
        SQL = SQL & DevNombreSQL(RS!Descripcion) & "','" & RS!Ctaprevista & "',"
       
  
        'Detalla
        If IsNull(RS!Fecha) Then
            SQL = SQL & "0,'" & Format(Now, FormatoFecha) & "'"
        Else
            SQL = SQL & TransformaComasPuntos(DBLet(RS!Importe, "N")) & ",'" & Format(RS!Fecha, FormatoFecha) & "'"
        End If
        SQL = SQL & ")"
        
        'Siguiente
        RS.MoveNext
            
        
    Wend
    RS.Close
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'QUITO LA COMA
        SQL = RC & SQL
        Conn.Execute SQL
    End If
        
        
    If NumRegElim = 0 Then
        MsgBox "Ningun dato devuelto", vbExclamation
        Exit Function
    End If
    
    
    'Updateo la cuenta bancaria
    RC = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY 1"
    RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
        SQL = SQL & RS!texto3 & "|"
        RS.MoveNext
    Wend
    RS.Close
    
    While SQL <> ""
        NumRegElim = InStr(1, SQL, "|")
        If NumRegElim = 0 Then
            SQL = ""
        Else
            RC = Mid(SQL, 1, NumRegElim - 1)
            SQL = Mid(SQL, NumRegElim + 1)
            
            RC = "Select codmacta,nommacta from cuentas where codmacta='" & RC & "'"
            RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                RC = "UPDATE usuarios.ztesoreriacomun SET texto4='" & DevNombreSQL(RS!Nommacta) & "' WHERE codusu =" & vUsu.Codigo & " AND texto3='" & RS!codmacta & "'"
                Conn.Execute RC
            End If
            RS.Close
        End If
    Wend
    ListadoGastosFijos = True
    Exit Function
EListadoGF:
    MuestraError Err.Number, Err.Description
End Function






'Listadoas aseguradoas
Private Function AvisosAseguradora() As Boolean



    On Error GoTo EListAsegEfectos
    AvisosAseguradora = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    'feccomunica,fecprorroga,fecsiniestro
    SQL = ""
    If Me.optAsegAvisos(0).Value Then
        Cad = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        Cad = "fecprorroga"
    Else
        Cad = "fecsiniestro"
    End If
    RC = CampoABD(Text3(21), "F", Cad, True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", Cad, False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Significa que no ha puesto fechas
    If InStr(1, SQL, Cad) = 0 Then SQL = SQL & " AND " & Cad & ">='1900-01-01'"
    
    'ORDENACION
    If Me.optAsegAvisos(0).Value Then
        RC = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        RC = "fecprorroga"
    Else
        RC = "fecsiniestro"
    End If
    
    Cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,numpoliz"
    Cad = Cad & ",credicon," & RC & " LaFecha" 'alias
    Cad = Cad & "  FROM scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa "
    If SQL <> "" Then Cad = Cad & SQL
    
    
    
    

    Cad = Cad & " ORDER BY " & RC & ","
    'ORDENACION 2� nivel
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & RC
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,faviso,fvto,impvto,disponible,vencida
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,fecha3,importe1,importe2,opcion) VALUES "
    RC = ""
    
    While Not miRsAux.EOF
        If Len(RC) > 500 Then
            RC = Mid(RC, 2)
            Conn.Execute Cad & RC
            RC = ""
        End If
        CONT = CONT + 1
        SQL = ", (" & vUsu.Codigo & "," & CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
        SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "'"
        SQL = SQL & ",'" & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"  'texto4
        'Fecha fac
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha aviso
        SQL = SQL & CampoBD_A_SQL(miRsAux!lafecha, "F", True) & ","
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecvenci, "F", True)
        
        SQL = SQL & "," & TransformaComasPuntos(CStr(miRsAux!impvenci))
        SQL = SQL & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!Gastos, "N")))
        'Devuelto
        SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        RC = RC & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If RC <> "" Then
        RC = Mid(RC, 2)
        Conn.Execute Cad & RC
    End If
    
    
    If CONT > 0 Then
        AvisosAseguradora = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "Avisos aseguradoras"
End Function



'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'       Compensaciones Cliente. Abonos vs Cobros
'
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Sub PonerVtosCompensacionCliente()
Dim IT


    lwCompenCli.ListItems.Clear
    Me.txtimpNoEdit(0).Tag = 0
    Me.txtimpNoEdit(1).Tag = 0
    Me.txtimpNoEdit(0).Text = ""
    Me.txtimpNoEdit(1).Text = ""
    If Me.txtCta(17).Text = "" Then Exit Sub
    Set Me.lwCompenCli.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    Cad = "Select scobro.*,nomforpa from scobro,sforpa where scobro.codforpa=sforpa.codforpa "
    Cad = Cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    Cad = Cad & " AND (transfer =0 or transfer is null) and codrem is null"
    Cad = Cad & " and estacaja=0 and recedocu=0"
    Cad = Cad & " ORDER BY fecvenci"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCompenCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!codfaccl, "000000")
        IT.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!fecvenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!impvenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Importe > 0 Then
            IT.SubItems(6) = Format(Importe, FormatoImporte)
            IT.SubItems(7) = " "
        Else
            IT.SubItems(6) = " "
            IT.SubItems(7) = Format(-Importe, FormatoImporte)
        End If
        IT.Tag = Abs(Importe)  'siempre valor absoluto
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub RealizarCompensacionAbonosClientes()
Dim Borras As Boolean
    
    If BloqueoManual(True, "COMPEABONO", "1") Then

        Cad = DevuelveDesdeBD("max(codigo)", "scompenclicab", "1", "1")
        If Cad = "" Then Cad = "0"
        CONT = Val(Cad) + 1 'ID de la operacion
        
        Cad = "INSERT INTO scompenclicab(codigo,fecha,login,PC,codmacta,nommacta) VALUES (" & CONT
        Cad = Cad & ",now(),'" & DevNombreSQL(vUsu.Login) & "','" & DevNombreSQL(vUsu.PC)
        Cad = Cad & "','" & txtCta(17).Text & "','" & DevNombreSQL(DtxtCta(17).Text) & "')"
        
        Set miRsAux = New ADODB.Recordset
        Borras = True
        If Ejecuta(Cad) Then
            
            Borras = Not RealizarProcesoCompensacionAbonos
        
        End If


        Set miRsAux = Nothing
        If Borras Then
            Conn.Execute "DELETE FROM scompenclilin WHERE codigo = " & CONT
            Conn.Execute "DELETE FROM scompenclicab WHERE codigo = " & CONT
            
        End If

        'Desbloquamos proceso
        BloqueoManual False, "COMPEABONO", ""
        DevfrmCCtas = ""
        
        PonerVtosCompensacionCliente   'Volvemos a cargar los vencimientos
        
        'El nombre del report
        CadenaDesdeOtroForm = Me.Tag
        Me.Tag = ""
        If Not Borras Then
            ImprimiCompensacion CONT
            
        
        End If
        
        Set miRsAux = Nothing
    Else
        MsgBox "Proceso bloqueado", vbExclamation
    End If

End Sub



Private Sub ImprimiCompensacion(CodigoCompensacion As Long)

    On Error GoTo EImprimiCompensacion
        
        'CadenaDesdeOtroForm:  lleva el nombre del report
        
        
        'Ha ido bien. Imprimiremos la hoja por si quiere crear PDF
        Conn.Execute "DELETE FROM Usuarios.ztmpfaclin WHERE codusu =" & vUsu.Codigo
        Conn.Execute "Delete from Usuarios.zTesoreriaComun where codusu = " & vUsu.Codigo
        Conn.Execute "Delete from Usuarios.z347carta where codusu = " & vUsu.Codigo
        
        'insert into `ztmpfaclin` (`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,
        '`Imponible`,`IVA`,`ImpIVA`,`Total`,`retencion`,`TipoIva`)
        SQL = "INSERT INTO usuarios.ztmpfaclin(`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,`Imponible`,`ImpIVA`,`retencion`,`Total`,`IVA`,TipoIva)"
        SQL = SQL & "select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,"
        SQL = SQL & "concat(numserie,right(concat(""000000"",codfaccl),8)) fecha,date_format(fecfaccl,'%d/%m/%Y') ffaccl,"
        SQL = SQL & "scompenclilin.codmacta,if (nommacta is null,nomclien,nommacta) nomcli,"
        SQL = SQL & "date_format(fecvenci,'%d/%m/%Y') venci,impvenci,gastos,impcobro,"
        SQL = SQL & "impvenci + coalesce(gastos,0) + coalesce(impcobro,0)  tot_al"
        SQL = SQL & ",if(fecultco is null,null,date_format(fecultco,'%d/&m')) fecco ,destino"
        SQL = SQL & " From (scompenclilin left join cuentas on scompenclilin.codmacta=cuentas.codmacta)"
        SQL = SQL & ",(SELECT @rownum:=0) r WHERE codigo=" & CONT & " order by destino desc,numserie,codfaccl"
        Conn.Execute SQL
            
        
            
        
   
    
    
        
    
    
    
    
        'Datos carta
        'Datos basicos de la empresa para la carta
        Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, "
        Cad = Cad & "parrafo1, parrafo2, contacto, despedida,saludos,parrafo3, parrafo4, parrafo5, Asunto, Referencia)"
        Cad = Cad & " VALUES (" & vUsu.Codigo & ", "
        
        'Estos datos ya veremos com, y cuadno los relleno
        Set miRsAux = New ADODB.Recordset
        SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia,contacto from empresa2"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Paarafo1 Parrafo2 contacto
        SQL = "'','',''"
        'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
        SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'," & SQL
        If Not miRsAux.EOF Then
            SQL = ""
            For I = 1 To 6
                SQL = SQL & DBLet(miRsAux.Fields(I), "T") & " "
            Next I
            SQL = Trim(SQL)
            SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
            SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"

            'Contaccto
            SQL = SQL & ",NULL,NULL,'" & DevNombreSQL(DBLet(miRsAux!contacto)) & "' "
        End If
        miRsAux.Close
      
        Cad = Cad & SQL
        Cad = Cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
        
        Conn.Execute Cad
        
        
        'Datos CLIENTE
         ', texto3, texto4, texto5,texto6
        Cad = DevuelveDesdeBD("codmacta", "scompenclicab", "codigo", CStr(CONT))
        Cad = "Select nommacta,razosoci,dirdatos,codposta,despobla,desprovi from cuentas where codmacta ='" & Cad & "'"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        'NO PUEDE SER EOF
        Cad = miRsAux!Nommacta
        If Not IsNull(miRsAux!razosoci) Then Cad = miRsAux!razosoci
        Cad = "'" & DevNombreSQL(Cad) & "'"
        'Direccion
        Cad = Cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!dirdatos))) & "'"
        'Poblacion
        SQL = DBLet(miRsAux!codposta)
        If SQL <> "" Then SQL = SQL & " - "
        SQL = SQL & DevNombreSQL(CStr(DBLet(miRsAux!despobla)))
        Cad = Cad & ",'" & SQL & "'"
        'Provincia
        Cad = Cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!desprovi))) & "'"
        miRsAux.Close
        

        
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,texto5,texto6, observa1, "
        SQL = SQL & "importe1, importe2, fecha1, fecha2, fecha3, observa2, opcion)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'',''," & Cad
        
        'select Numfac,fecha from usuarios.ztmpfaclin where tipoiva=1 and codusu=2200
        Importe = 0
        RC = "NIF"   'RC = "fecha"   La fecha de VTo esta en el campo: NIF
        Cad = DevuelveDesdeBD("numfac", "Usuarios.ztmpfaclin", "codusu =" & vUsu.Codigo & " AND tipoiva", "1", "N", RC)
        If Cad = "" Then
            'Significa que la compesacion ha sido total, no quedaba resultante
            
        Else
            Cad = "Quedando el resultado en el vencimiento: " & Cad & " de " & Format(RC, "dd/mm/yyyy")
            Importe = 1
        End If
        
        If Importe > 0 Then
            RC = "select sum(impvenci + coalesce(gastos,0) + coalesce(impcobro,0)) from  scompenclilin  WHERE codigo=" & CONT
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = "0"
            If Not miRsAux.EOF Then Importe = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
        Else
            RC = "0"
        End If
        
        'observa 1 texto 6 e importe1
        SQL = SQL & ",'" & Cad & "'," & TransformaComasPuntos(CStr(Importe))
        
        
        'importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        For I = 1 To 6
            SQL = SQL & ",NULL"
        Next
        SQL = SQL & ")"
        Conn.Execute SQL
        
        With frmImprimir
                .OtrosParametros = ""
                .NumeroParametros = 0
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                
                .Opcion = 91
                .Show vbModal
            End With



Exit Sub
EImprimiCompensacion:
    MuestraError Err.Number, Err.Description
End Sub

Private Function RealizarProcesoCompensacionAbonos() As Boolean
Dim Destino As Byte
Dim J As Integer

    'NO USAR CONT

    RealizarProcesoCompensacionAbonos = False











    'Vamos a seleccionar los vtos
    '(numserie,codfaccl,fecfaccl,numorden)
    'EN SQL
    SQLVtosSeleccionadosCompensacion NumRegElim, False    'todos  -> Numregelim tendr el destino
    
    'Metemos los campos en el la tabla de lineas
    ' Esto guarda el valor en CAD
    FijaCadenaSQLCobrosCompen
    
    
    'Texto compensacion
    DevfrmCCtas = ""
    
    RC = "Select " & Cad & " FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & SQL & ")"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error. EOF vencimientos devueltos ", vbExclamation
        Exit Function
    End If
    
    
    I = 0
    
    While Not miRsAux.EOF
        I = I + 1
        BACKUP_Tabla miRsAux, RC
        'Quito los parentesis
        RC = Mid(RC, 1, Len(RC) - 1)
        RC = Mid(RC, 2)
        
        Destino = 0
        If miRsAux!NUmSerie = Me.lwCompenCli.ListItems(NumRegElim).Text Then
            If miRsAux!codfaccl = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!fecfaccl, "dd/mm/yyyy") = Me.lwCompenCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!numorden = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO scompenclilin (codigo,linea,destino," & Cad & ") VALUES (" & CONT & "," & I & "," & Destino & "," & RC & ")"
        Conn.Execute RC
        
        'Para las observaciones de despues
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!impvenci
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Destino = 0 Then 'El destino
            DevfrmCCtas = DevfrmCCtas & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000") & " " & Format(miRsAux!fecfaccl, "dd/mm/yy")
            DevfrmCCtas = DevfrmCCtas & " Vto:" & Format(miRsAux!fecvenci, "dd/mm/yy") & " " & Importe
            DevfrmCCtas = DevfrmCCtas & "|"
        Else
            'El DESTINO siempre ira en la primera observacion del texto
            RC = "Importe anterior vto: " & Importe
            DevfrmCCtas = RC & "|" & DevfrmCCtas
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Acutalizaremos el VTO destino
    
    Conn.BeginTrans
        'BORRAREMOS LOS VENCIMIENTOS QUE NO SEAN DESTINO a no ser que el importe restante sea 0
        Destino = 1
        If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then Destino = 0
        SQLVtosSeleccionadosCompensacion 0, Destino = 1  'sin o con el destino
        RC = "DELETE FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & SQL & ")"
        
        'Para saber si ha ido bien
        Destino = 0    '0 mal,1 bien
        If Ejecuta(RC) Then
            If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then
                Destino = 1
            Else
                'Updatearemos los campos csb del vto restante. A partir del segundo
                'La variable CadenaDesdeOtroForm  tiene los que vamos a actualizar
                
                Cad = ""
                J = 0
                SQL = ""
                
                Do
                    I = InStr(1, DevfrmCCtas, "|")
                    If I = 0 Then
                        DevfrmCCtas = ""
                    Else
                        RC = Mid(DevfrmCCtas, 1, I - 1)
                        If Len(RC) > 40 Then RC = Mid(RC, 1, 40)
                        DevfrmCCtas = Mid(DevfrmCCtas, I + 1)
                        J = J + 1
                        'Antes desde aqui cogia el campo
                        'Ahora desde CadenaDesdeOtroForm que tiene los campos libres
                        'Cad = RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
                        Cad = RecuperaValor(CadenaDesdeOtroForm, J)
                        SQL = SQL & ", " & Cad & " = '" & DevNombreSQL(RC) & "'"
                
                    End If
                Loop Until DevfrmCCtas = ""
                Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
                RC = "gastos=null, impcobro=null,fecultco=null,impvenci=" & TransformaComasPuntos(CStr(Importe))
                SQL = RC & SQL
                SQL = "UPDATE scobro SET " & SQL
                'WHERE
                RC = ""
                For J = 1 To Me.lwCompenCli.ListItems.Count
                    If Me.lwCompenCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwCompenCli.ListItems(J).Text
                        RC = RC & "' AND codfaccl = " & Val(Me.lwCompenCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfaccl = '" & Format(Me.lwCompenCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwCompenCli.ListItems(J).SubItems(3))
                        Exit For
                    End If
                Next
                If RC <> "" Then
                    Cad = SQL & " WHERE " & RC
                    If Ejecuta(Cad) Then Destino = 1
                Else
                    MsgBox "No encontrado destino", vbExclamation
                    
                End If
            End If
        End If
        If Destino = 1 Then
            Conn.CommitTrans
            RealizarProcesoCompensacionAbonos = True
        Else
            Conn.RollbackTrans
        End If
        
End Function

Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
        
            Insertar = True
            If Me.lwCompenCli.ListItems(I).Bold Then
                RegistroDestino = I
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCompenCli.ListItems(I).Text & "'," & lwCompenCli.ListItems(I).SubItems(1)
                SQL = SQL & ",'" & Format(lwCompenCli.ListItems(I).SubItems(2), FormatoFecha) & "'," & lwCompenCli.ListItems(I).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    Cad = "NUmSerie , codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, ctabanc1,"
    Cad = Cad & "codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, emitdocum, "
    Cad = Cad & "recedocu, contdocu, text33csb, text41csb, text42csb, text43csb, text51csb, text52csb,"
    Cad = Cad & "text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb,"
    Cad = Cad & "text82csb, text83csb, ultimareclamacion, agente, departamento, tiporem, CodRem, AnyoRem,"
    Cad = Cad & "siturem, Gastos, Devuelto, situacionjuri, noremesar, obs, transfer, estacaja, referencia,"
    Cad = Cad & "reftalonpag, nomclien, domclien, pobclien, cpclien, proclien, referencia1, referencia2,"
    Cad = Cad & "feccomunica, fecprorroga, fecsiniestro"
    
End Sub


'******************************************************************************
'******************************************************************************
'
'******************************************************************************
'******************************************************************************



Private Function ComunicaDatosSeguro_() As Boolean
Dim K As Integer

    ComunicaDatosSeguro_ = False
    
   
    NumRegElim = 0
    
    For K = 1 To Me.ListView3.ListItems.Count
        If Me.ListView3.ListItems(K).Checked Then
            DatosSeguroUnaEmpresa CInt(ListView3.ListItems(K).Tag)
      
            SQL = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
            If SQL <> "" Then NumRegElim = Val(SQL)
        End If
    Next
    
    
    
    If NumRegElim > 0 Then
        SQL = "DELETE from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " AND importe1<=0"
        
        
        
    
    
        '   Conn.Execute SQL
        SQL = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
        If SQL <> "" Then
            NumRegElim = Val(SQL)
        Else
            NumRegElim = 0
        End If
        
        
        ComunicaDatosSeguro_ = NumRegElim > 0
        If NumRegElim > 0 Then
            SQL = "Select texto5 from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not miRsAux.EOF
                SQL = miRsAux!texto5
                If SQL = "" Then
                    SQL = "ESPA�A"
                Else
                    If InStr(1, SQL, " ") > 0 Then
                        SQL = Mid(SQL, 3)
                    Else
                        SQL = "" 'no updateamos
                    End If
                End If
                If SQL <> "" Then
                    SQL = "UPDATE Usuarios.ztesoreriacomun set texto5='" & DevNombreSQL(SQL) & "' WHERE codusu ="
                    SQL = SQL & vUsu.Codigo & " AND texto5='" & DevNombreSQL(miRsAux!texto5) & "'"
                    Conn.Execute SQL
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
    End If
End Function

Private Sub DatosSeguroUnaEmpresa(NumConta As Integer)

    
    'select numpoliz,nifdatos,numserie,codfaccl,nommacta,impvenci,gastos,impcobro,credicon from scobro,cuentas where
    'scobro.codmacta = cuentas.codmacta     fecbajcre
    
    'JUlio2013
    'Para fontenas iran por PAIS
    'a�adiremos en text05 el pais
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
    SQL = SQL & " importe1,  importe2,texto5) "
    
'    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
'
'    RC = RC & "impvenci + if(gastos is null,0,gastos) - if( impcobro is null,0,impcobro) ,credicon"
'    RC = RC & "  from conta" & NumConta & ".scobro,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
'    RC = RC & " WHERE scobro.codmacta=cuentas.codmacta  and numpoliz<>''  and "
'    RC = RC & " (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
'
    
    'ENERO 2013.
    'Despues de hablar con BERNIA, en este listado salen
    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
    
    RC = RC & " totfaccl ,credicon,if(pais is null,'',pais)"    'JUL13 a�adimos PAIS
    RC = RC & " from conta" & NumConta & ".cabfact,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
    RC = RC & " WHERE cabfact.codmacta=cuentas.codmacta  and numpoliz<>''  and "
    
    RC = RC & " (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
    
    'Contemplamos facturas desde la fecha de concesion
    RC = RC & " and fecfaccl>= fecconce"
    
    'D/H fecha factura
    If Me.Text3(34).Text <> "" Then RC = RC & " AND fecfaccl >='" & Format(Text3(34).Text, FormatoFecha) & "'"
    If Me.Text3(35).Text <> "" Then RC = RC & " AND fecfaccl <='" & Format(Text3(35).Text, FormatoFecha) & "'"
    
    
    
    
    
    
    SQL = SQL & RC
    Conn.Execute SQL
End Sub


Private Function GeneraDatosFrasAsegurados() As Boolean
Dim NumConta As Byte

    NumConta = CByte(vEmpresa.codempre)
    GeneraDatosFrasAsegurados = False

    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
    SQL = SQL & " importe1,  importe2,fecha1,fecha2) "
    
    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
    
    RC = RC & "impvenci + if(gastos is null,0,gastos) - if( impcobro is null,0,impcobro) ,if (credicon is null,0,credicon)"
    RC = RC & ",fecfaccl,fecvenci"
    RC = RC & "  from conta" & NumConta & ".scobro,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
    RC = RC & " WHERE scobro.codmacta=cuentas.codmacta  "
    
    If Me.chkVarios(0).Value = 1 Then
        'SOLO asegudaros
        RC = RC & " and numpoliz<>''  and (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
    End If
    'D/H fecha factura
    If Me.Text3(34).Text <> "" Then RC = RC & " AND fecfaccl >='" & Format(Text3(34).Text, FormatoFecha) & "'"
    If Me.Text3(35).Text <> "" Then RC = RC & " AND fecfaccl <='" & Format(Text3(35).Text, FormatoFecha) & "'"
    
    
    SQL = SQL & RC
    Conn.Execute SQL

    
    
    'Borramos importe cero

    SQL = "DELETE from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " AND importe1<=0"
    Conn.Execute SQL
    
    SQL = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
    If SQL <> "" Then
        NumRegElim = Val(SQL)
    Else
        NumRegElim = 0
    End If
    GeneraDatosFrasAsegurados = NumRegElim > 0



End Function

'****************************************************************************************
'****************************************************************************************
'
'       NORMA 57
'
'****************************************************************************************
'****************************************************************************************
Private Function procesarficheronorma57() As Boolean
Dim Estado As Byte  '0  esperando cabcerea
                    '1  esperando pie (leyendo lineas)
    
    On Error GoTo eprocesarficheronorma57
    
    
    'insert into tmpconext(codusu,cta,fechaent,Pos)
    Conn.Execute "DELETE FROM tmpconext WHERE codusu = " & vUsu.Codigo
    procesarficheronorma57 = False
    I = FreeFile
    Open cd1.FileName For Input As #I
    SQL = ""
    Estado = 0
    Importe = 0
    TotalRegistros = 0
    While Not EOF(I)
            Line Input #I, SQL
            RC = Mid(SQL, 1, 4)
            Select Case Estado
            Case 0
                'Para saber que el fichero tiene el formato correcto
                If RC = "0270" Then
                        Estado = 1
                        'Voy a buscar si hay un banco
                        
                        RC = "select cuentas.codmacta,nommacta from ctabancaria,cuentas where ctabancaria.codmacta="
                        RC = RC & "cuentas.codmacta AND ctabancaria.entidad = " & Trim(Mid(SQL, 23, 4))
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        TotalRegistros = 0
                        While Not miRsAux.EOF
                            RC = miRsAux!codmacta & "|" & miRsAux!Nommacta & "|"
                            TotalRegistros = TotalRegistros + 1
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                        If TotalRegistros = 1 Then
                            Me.txtCtaBanc(5).Text = RecuperaValor(RC, 1)
                            Me.txtDescBanc(5).Text = RecuperaValor(RC, 2)
                        End If
                        TotalRegistros = 0
                End If
            Case 1
                If RC = "6070" Then
                    'Linea con recibo
                    'Ejemplo:
                    '   6070      46076147000130582263151014000000014067003059                      0000000516142
                    '                                  fecha       impot   socio                      fra      CC codigo de control del codigo de barra
                    'Fecha pago
                    RC = Mid(SQL, 31, 2) & "/" & Mid(SQL, 33, 2) & "/20" & Mid(SQL, 35, 2)
                    Fecha = CDate(RC)
                    'IMporte
                    RC = Mid(SQL, 37, 12)
                    Cad = CStr(CCur(Val(RC) / 100))
                    'FRA
                    RC = Mid(SQL, 77, 11)
                    CONT = Val(RC)
                    'Socio
                    RC = Val(Mid(SQL, 50, 6))
                        
                    'Insertamos en tmp
                    TotalRegistros = TotalRegistros + 1
                    SQL = "INSERT INTO tmpconext(codusu,cta,fechaent,Pos,TimporteD,linliapu) VALUES (" & vUsu.Codigo & ",'"
                    SQL = SQL & RC & "','" & Format(Fecha, FormatoFecha) & "'," & CONT & "," & TransformaComasPuntos(Cad) & "," & TotalRegistros & ")"
                    Conn.Execute SQL
                    
                    Importe = Importe + CCur(TransformaPuntosComas(Cad))
                ElseIf RC = "8070" Then
                    'OK. Final de linea.
                    '
                    'Comprobacion BASICA
                    '8070      46076147000 000010        000000028440
                    '                       vtos-2           importe
                    
                    RC = ""
                    
                    'numero registros
                    Cad = Val(Mid(SQL, 24, 5))
                    If Val(Cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. N� registros cero. " & SQL
                    Else
                        If Val(Cad) - 2 <> TotalRegistros Then RC = "Contador de registros incorrecto"
                    End If
                    'Suma importes
                    Cad = CStr(CCur(Mid(SQL, 37, 12) / 100))
                    
                    If CCur(Cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Suma importes cero. " & SQL
                    Else
                        If CCur(Cad) <> Importe Then RC = RC & vbCrLf & "Suma importes incorrecta"
                    End If
                    
                    
                   
                    
                    
                    If RC <> "" Then
                        Err.Raise 513, , RC
                    Else
                        Estado = 2
                    End If
                End If
            End Select
    Wend
    Close #I
    I = 0 'para que no vuelva a cerrar el fichero
    
    If Estado < 2 Then
        'Errores procesando fichero
        If Estado = 0 Then
            SQL = "No se encuetra la linea de inicio de declarante(6070)"
        Else
            SQL = "No se encuetra la linea de totales(8070)"
        End If

        MsgBox "Error procesando el fichero." & vbCrLf & SQL, vbExclamation
    Else
        espera 0.5
        procesarficheronorma57 = True
    End If
eprocesarficheronorma57:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    If I > 0 Then Close #I
End Function


Private Function BuscarVtosNorma57() As Boolean

    BuscarVtosNorma57 = False
    
    Set miRsAux = New ADODB.Recordset
    
    'Dependiendo del parametro....
    If vParam.Norma57 = 1 Then
        'ESCALONA.
        'Viene el socio y el numero de factura e importe.
        'Habra que buscar
        BuscarVtosNorma57 = VtosNorma57Escalona

    Else
        MsgBox "En desarrollo", vbExclamation
    End If
    
    Set miRsAux = Nothing
End Function

Private Function VtosNorma57Escalona() As Boolean
Dim RN As ADODB.Recordset
Dim Fin As Boolean
Dim NoEncontrado As Byte
Dim AlgunVtoNoEncontrado As Boolean
On Error GoTo eVtosNorma57Escalona
    
    VtosNorma57Escalona = False
    Set RN = New ADODB.Recordset
    SQL = "select * from tmpconext WHERE codusu =" & vUsu.Codigo & " order by cta,pos "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AlgunVtoNoEncontrado = False
    While Not miRsAux.EOF
        'Vto a vto
        'If miRsAux!Linliapu = 9 Then Stop
        RC = RellenaCodigoCuenta("430." & miRsAux!Cta)
        SQL = "Select * from scobro where codmacta = '" & RC & "' AND codfaccl =" & miRsAux!Pos & " and impvenci>0"
        RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = 0
        If RN.EOF Then
            Cad = "NO encontrado"
            NoEncontrado = 2
        Else
            'OK encontrado.
            Fin = False
            I = 0
            NoEncontrado = 1
            Cad = ""
            
            While Not Fin
            
                I = I + 1
                
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                If Not Fin Then
                    RN.MoveNext
                    If RN.EOF Then Fin = True
                End If
            Wend
        End If
        RN.Close
        SQL = "UPDATE tmpconext SET "
        If CONT = 1 Then
            'OK este es el vto
            'NO hacemos nada. Updateamos los campos de la tmp
            'para buscar despues
            'numdiari numorden       numdocum=fecfaccl     ccost numserie
            SQL = SQL & " nomdocum ='" & Format(Fecha, FormatoFecha)
            SQL = SQL & "', ccost ='" & DevfrmCCtas
            SQL = SQL & "', numdiari = " & I
            SQL = SQL & ", contra = '" & RC & "'"
        Else
            If I > 1 Then Cad = "(+1) " & Cad
            SQL = SQL & " numasien=  " & NoEncontrado  'para vtos no encontrados o erroneos
            SQL = SQL & ", ampconce ='" & DevNombreSQL(Cad) & "'"
            If NoEncontrado = 2 Then AlgunVtoNoEncontrado = True
        End If
        SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND linliapu = " & miRsAux!Linliapu
        Conn.Execute SQL
            
 
        
        'Sig
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    
    
    If AlgunVtoNoEncontrado Then
        'Lo buscamos al reves
        espera 0.5
        SQL = "select * from  tmpconext  WHERE codusu =" & vUsu.Codigo & " AND numasien=2"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Miguel angel
            'Puede que en algunos recibos las posciones del fichero vengan cambiadas
            'Donde era la factura es la cta y al reves
            RC = RellenaCodigoCuenta("430." & miRsAux!Pos)
            SQL = "Select * from scobro where codmacta = '" & RC & "' AND codfaccl =" & Val(miRsAux!Cta) & " and impvenci>0"
            RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
        
                'OK encontrado.
                Fin = False
                CONT = 0
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                
            
            
                'OK este es el vto
                'NO hacemos nada. Updateamos los campos de la tmp
                'para buscar despues
                'numdiari numorden       numdocum=fecfaccl     ccost numserie
                If CONT = 1 Then
                    SQL = SQL & " nomdocum ='" & Format(Fecha, FormatoFecha)
                    SQL = SQL & "', ccost ='" & DevfrmCCtas
                    SQL = SQL & "', numdiari = " & I
                    SQL = SQL & ", contra = '" & RC & "'"
                    SQL = "UPDATE tmpconext SET "
                End If
            End If
            RN.Close
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    VtosNorma57Escalona = True
eVtosNorma57Escalona:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, "Buscando vtos Escalona"
    Set RN = Nothing
End Function


Private Sub Norma57VencimientoEncontradoEsCorrecto(ByRef Rss As ADODB.Recordset, ByRef Final As Boolean)
        
        'Ha encontrado el vencimiento. Falta ver si no esta en remesa....
        If Not IsNull(Rss!CodRem) Then
            Cad = "En la remesa " & Rss!CodRem
        
        Else
            If Not IsNull(Rss!transfer) Then
                Cad = "Transferencia " & Rss!transfer
            Else
                Importe = Rss!impvenci + DBLet(Rss!Gastos, "N") - DBLet(Rss!impcobro, "N")
                If Importe <> miRsAux!timported Then
                    'Importe distinto
                    'Veamos si es que esta
                    Cad = "Importe distinto"
                Else
                    'OK. Misma factura, socio, importe. SAlimos ya poniendo ""
                    Fecha = Rss!fecfaccl
                    DevfrmCCtas = Rss!NUmSerie
                    I = Rss!numorden
                    Cad = ""
                    Final = True
                    CONT = 1
                End If
            End If
        End If
End Sub

Private Sub CargaLWNorma57(Correctos As Boolean)
Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    If Correctos Then
        SQL = "select tmpconext.*,nommacta from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " and numasien=0 order by  ccost,pos  "
    Else
        SQL = "select * from tmpconext WHERE codusu = " & vUsu.Codigo & " and numasien > 0 order by cta,pos "
    End If
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Correctos Then
            Set IT = Me.lwNorma57Importar(0).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!CCost
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!Nomdocum, "dd/mm/yyyy")
            IT.SubItems(3) = miRsAux!Linliapu
            If IsNull(miRsAux!Nommacta) Then
                SQL = "ERRROR GRAVE"
            Else
                SQL = miRsAux!Nommacta
            End If
            IT.SubItems(4) = SQL
            IT.SubItems(5) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(6) = Format(miRsAux!fechaent, "dd/mm/yyyy")
            IT.Checked = True
        Else
            'ERRORES
            Set IT = Me.lwNorma57Importar(1).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!Cta
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(3) = miRsAux!ampconce
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub


Private Sub LimpiarDelProceso()
    lwNorma57Importar(0).ListItems.Clear
    lwNorma57Importar(1).ListItems.Clear
    Me.txtCtaBanc(5).Text = ""
    Me.txtDescBanc(5).Text = ""
End Sub



'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
' Cobros parciales
Private Sub HabilitarCobrosParciales()
    On Error GoTo eHabilitarCobrosParciales
    Me.cmdCobrosAgenLin.Enabled = False
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from scobrolin where codfaccl=-1", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miRsAux.Close
    Me.cmdCobrosAgenLin.Enabled = True

eHabilitarCobrosParciales:
    Err.Clear
    Set miRsAux = Nothing
End Sub

Private Function DesdeHastaAgenteCobrosParciales() As String

    'SQL desde/hasta
    SQL = ""
    DevfrmCCtas = "" 'VISREPORT. Desde hasta
    RC = ""
    If Text3(36).Text <> "" Then
        SQL = SQL & " AND fecha >= '" & Format(Text3(36).Text, FormatoFecha) & "'"
        RC = RC & " desde " & Text3(36).Text
    End If
    If Text3(37).Text <> "" Then
        SQL = SQL & " AND fecha <= '" & Format(Text3(37).Text, FormatoFecha) & "'"
        RC = RC & " hasta " & Text3(37).Text
    End If
    If RC <> "" Then DevfrmCCtas = "Fecha cobro " & RC
    RC = ""
    If Me.txtAgente(6).Text <> "" Then
        SQL = SQL & " AND codagent >=" & txtAgente(6).Text
        RC = RC & " desde " & Me.txtAgente(6).Text & " " & Me.txtDescAgente(6).Text
    End If
    If Me.txtAgente(7).Text <> "" Then
        SQL = SQL & " AND codagent <=" & txtAgente(7).Text
        RC = RC & " hasta " & Me.txtAgente(7).Text & " " & Me.txtDescAgente(7).Text
    End If
    If RC <> "" Then DevfrmCCtas = Trim(DevfrmCCtas & "    Agentes " & RC)
    
    
    
    If SQL <> "" Then SQL = Mid(SQL, 5) 'quito el primer AND
    DesdeHastaAgenteCobrosParciales = SQL


End Function


Private Function GenerarDatosListadoCobrosParcialesAgente() As Boolean

On Error GoTo eGenerarDatosListadoCobrosParcialesAgente

    GenerarDatosListadoCobrosParcialesAgente = False
    Set miRsAux = New ADODB.Recordset
    
    Conn.Execute "DELETE FROM usuarios.zpendientes WHERE codusu =" & vUsu.Codigo
    
    
    RC = DesdeHastaAgenteCobrosParciales
    SQL = RC
    
    'Comprobacion. QUE todos los datos en linea de cobros tienen recibo en scobro
    '------------------------------------------------------------------------------------
    Label3(50).Caption = "Comprobaciones "
    RC = ""
    'Vtros
    Cad = "Select * from scobrolin where not (numserie,codfaccl,fecfaccl,numorden) IN ("
    Cad = Cad & "SELECT numserie,codfaccl,fecfaccl,numorden FROM scobro )"
    If SQL <> "" Then Cad = Cad & " AND " & SQL
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & "- " & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "000000") & " " & miRsAux!fecfaccl & " (" & miRsAux!numorden & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Cad <> "" Then RC = RC & "Vtos NO existen. " & vbCrLf & Cad & vbCrLf & vbCrLf
    'Agentes
    Cad = "Select * from scobrolin where not (codagent) IN ("
    Cad = Cad & "SELECT codigo FROM agentes )"
    If SQL <> "" Then Cad = Cad & " AND " & SQL
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & "- " & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "000000") & " " & miRsAux!fecfaccl & " (" & miRsAux!numorden & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Cad <> "" Then RC = RC & "Agentes NO existen. " & vbCrLf & Cad & vbCrLf & vbCrLf
    'Foramas de pago
    Cad = "Select * from scobrolin where not (codagent) IN ("
    Cad = Cad & "SELECT codigo FROM agentes )"
    If SQL <> "" Then Cad = Cad & " AND " & SQL
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & "- " & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "000000") & " " & miRsAux!fecfaccl & " (" & miRsAux!numorden & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Cad <> "" Then RC = RC & "Formas de pago NO existen. " & vbCrLf & Cad & vbCrLf & vbCrLf
    
    If RC <> "" Then Err.Raise 513, , RC
    
    
    
    'Insertamos
    Cad = "INSERT INTO usuarios.zpendientes(codusu,serie_cta,factura,fecha,numorden,nomdirec,fecVto,importe,codforpa,   observa)"
    Cad = Cad & "Select " & vUsu.Codigo & ",numserie,codfaccl,fecfaccl,numorden*100 + id,codagent,fecha,importe,codforpa,observa FROM scobrolin"
    If SQL <> "" Then Cad = Cad & " WHERE " & SQL
    Conn.Execute Cad
    
    espera 0.5
    'Cliente
    Cad = "Select " & vUsu.Codigo & ",numserie,codfaccl,fecfaccl FROM scobrolin"
    If SQL <> "" Then Cad = Cad & " WHERE " & SQL
    Cad = Cad & " GROUP BY 1,2,3,4"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = "apudirec='S' AND scobro.codmacta = cuentas.codmacta AND numserie='" & miRsAux!NUmSerie & "' "
        Cad = Cad & " AND fecfaccl='" & Format(miRsAux!fecfaccl, FormatoFecha) & "' AND codfaccl"
        RC = "cuentas.codmacta"
        Cad = DevuelveDesdeBD("cuentas.nommacta", "scobro,cuentas", Cad, miRsAux!codfaccl, "N", RC)
        If Cad = "" Then
            Cad = "Error VTO cuenta cliente." & vbCrLf & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "000000") & " " & miRsAux!fecfaccl & " (" & miRsAux!numorden & ")"
            Err.Raise 513, , RC
        Else
            NombreSQL Cad
            Cad = "UPDATE usuarios.zpendientes SET nombre='" & Cad & "',codmacta ='" & RC & "'"
            Cad = Cad & " WHERE serie_cta='" & miRsAux!NUmSerie & "' AND factura=" & miRsAux!codfaccl
            Cad = Cad & " AND fecha='" & Format(miRsAux!fecfaccl, FormatoFecha) & "'"
            Conn.Execute Cad
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'AGENTES
    Cad = "Select * from usuarios.zpendientes WHERE codusu = " & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    While Not miRsAux.EOF
        Cad = DevuelveDesdeBD("nombre", "agentes", "codigo", miRsAux!nomdirec)
        If Cad = "" Then
            Cad = "******  ERROR leyendo agente"
        Else
            Cad = Format(miRsAux!nomdirec, "0000") & " " & Cad
        End If
        NombreSQL Cad
        RC = "UPDATE usuarios.zpendientes SET nomdirec='" & Cad & "' WHERE codusu =" & vUsu.Codigo & " AND nomdirec= '" & miRsAux!nomdirec & "'"
        Conn.Execute RC
        miRsAux.MoveNext
        CONT = CONT + 1
    Wend
    miRsAux.Close
    Cad = "Select distinct codforpa from usuarios.zpendientes WHERE codusu = " & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", miRsAux!codforpa)
        If Cad = "" Then Cad = "******  ERROR leyendo forma de pago"
        
        NombreSQL Cad
        RC = "UPDATE usuarios.zpendientes SET nomforpa='" & Cad & "' WHERE codusu =" & vUsu.Codigo & " AND codforpa= " & miRsAux!codforpa
        Conn.Execute RC
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    
    If CONT = 0 Then
        MsgBox "Ningun datos generado", vbExclamation
        Label3(50).Caption = ""
    Else
        GenerarDatosListadoCobrosParcialesAgente = True
    End If
eGenerarDatosListadoCobrosParcialesAgente:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function


Private Function RealizarProcesoUpdateCobrosAgente() As Boolean
Dim RCob As ADODB.Recordset
Dim vLog As cLOG

    Set vLog = New cLOG
    On Error GoTo eRealizarProcesoUpdateCobrosAgente
    RealizarProcesoUpdateCobrosAgente = False
    Set miRsAux = New ADODB.Recordset
    Set RCob = New ADODB.Recordset
    
    'tiene que sumar lo cobrado por factura y saldar el cobro
    RC = DesdeHastaAgenteCobrosParciales
    Cad = "select numserie,codfaccl,fecfaccl,numorden,sum(importe) as cobrado,count(*) as cuantos from scobrolin "
    If RC <> "" Then Cad = Cad & " WHERE " & RC
    Cad = Cad & " group by 1,2,3,4 "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label3(50).Caption = "Cobro " & miRsAux!NUmSerie & miRsAux!codfaccl
        Label3(50).Refresh
        RC = "numserie = '" & miRsAux!NUmSerie & "' AND codfaccl =" & miRsAux!codfaccl
        RC = RC & " AND fecfaccl = '" & Format(miRsAux!fecfaccl, FormatoFecha) & "' AND numorden=" & miRsAux!numorden
        Cad = "SELECT * from scobro where " & RC
        RCob.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RCob!codmacta, "T")
        Cad = Cad & " " & miRsAux!NUmSerie & miRsAux!codfaccl & " de " & Format(miRsAux!fecfaccl, "dd/mm/yyyy") & " -" & miRsAux!numorden
        Cad = Cad & vbCrLf & "�Recibo: " & RCob!impvenci
        If Not IsNull(RCob!Gastos) Then Cad = Cad & " Gastos:" & RCob!Gastos
        If Not IsNull(RCob!impcobro) Then Cad = Cad & " Ult cobro:" & RCob!impcobro
        Cad = Cad & vbCrLf & "� Cobrados agente: " & miRsAux!cobrado
        Cad = Cad & vbCrLf & "N� cobros: " & miRsAux!Cuantos
        
        
        Importe = DBLet(RCob!impcobro, "N") + miRsAux!cobrado
        
        If Importe < RCob!impvenci + DBLet(RCob!Gastos, "N") Then
            SQL = "UPDATE scobro  SET impcobro = " & TransformaComasPuntos(CStr(Importe))
            SQL = SQL & ", fecultco = '" & Format(Now, FormatoFecha) & "'"
        Else
            Importe = Importe - (RCob!impvenci + DBLet(RCob!Gastos, "N"))
            If Importe > 0 Then Cad = Cad & vbCrLf & "Diferencia postiva: " & Importe
            SQL = "DELETE FROM scobro "
        End If
        SQL = SQL & " WHERE " & RC
        Conn.Execute SQL
        vLog.Insertar 101, vUsu, Cad
        
        RCob.Close
        
        espera 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Eliminamos cobros parciales
    Label3(50).Caption = "Eliminando parciales"
    Label3(50).Refresh
    RC = DesdeHastaAgenteCobrosParciales
    Cad = "DELETE from scobrolin "
    If RC <> "" Then Cad = Cad & " WHERE " & RC
    Conn.Execute Cad
    
        
    
    RealizarProcesoUpdateCobrosAgente = True
    
eRealizarProcesoUpdateCobrosAgente:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set vLog = Nothing
    Label3(50).Caption = ""
    
End Function
