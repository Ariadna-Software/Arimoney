VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBanco 
   Caption         =   "Bancos propios"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "frmBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   26
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   28
      Tag             =   "Cedante|T|S|||ctabancaria|DocPagare|||"
      Text            =   "Text1"
      Top             =   7440
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   25
      Left            =   120
      MaxLength       =   40
      TabIndex        =   27
      Tag             =   "Cedante|T|S|||ctabancaria|DocTalon|||"
      Text            =   "Text1"
      Top             =   7440
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   24
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "C|T|S|||ctabancaria|CaixaConfirming|||"
      Text            =   "Text1"
      Top             =   2520
      Width           =   2445
   End
   Begin VB.CheckBox chkBanco 
      Caption         =   " .-"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Tag             =   "SEPA19.14|N|S|||ctabancaria|N1914GrabaNifDeudor|||"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   23
      Left            =   120
      MaxLength       =   3
      TabIndex        =   10
      Tag             =   "Cedante|T|S|||ctabancaria|Sufijo3414|||"
      Text            =   "Text1"
      Top             =   2520
      Width           =   1245
   End
   Begin VB.CheckBox chkBanco 
      Caption         =   " .-"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   70
      Tag             =   "G.transfer|N|S|||ctabancaria|GastTransDescontad|||"
      Top             =   6840
      Width           =   495
   End
   Begin VB.Frame FrameAnalitica 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4920
      TabIndex        =   69
      Top             =   6360
      Width           =   5535
   End
   Begin VB.CheckBox chkBanco 
      Caption         =   " .-"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Tag             =   "G.Rem.|N|S|||ctabancaria|GastRemDescontad|||"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "Remesas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   63
      Top             =   5280
      Width           =   10215
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Talon dias|N|S|0||ctabancaria|remesadiasmenor|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Talon dias|N|S|0||ctabancaria|remesadiasmayor|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Riesgo|N|S|0||ctabancaria|remesariesgo|#,##0.00||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   8760
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Riesgo|N|S|0||ctabancaria|remesamaximo|#,##0.00||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo menor"
         Height          =   255
         Index           =   22
         Left            =   5160
         TabIndex        =   67
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo mayor"
         Height          =   195
         Index           =   21
         Left            =   3000
         TabIndex        =   66
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Importe  riesgo"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe máximo "
         Height          =   195
         Index           =   18
         Left            =   7440
         TabIndex        =   64
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame FramePagares 
      Caption         =   "Pagarés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   60
      Top             =   4320
      Width           =   5055
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Talon dias|N|S|0||ctabancaria|pagaredias|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   16
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Riesgo|N|S|0||ctabancaria|pagareriesgo|#,##0.00||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   62
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Importe máximo "
         Height          =   195
         Index           =   16
         Left            =   2280
         TabIndex        =   61
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Talones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   57
      Top             =   4320
      Width           =   5055
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Riesgo|N|S|0||ctabancaria|talonriesgo|#,##0.00||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Talon dias|N|S|0||ctabancaria|talondias|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Importe máximo "
         Height          =   195
         Index           =   15
         Left            =   2400
         TabIndex        =   59
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   16
      Tag             =   "Cta. gastos|T|S|||ctabancaria|ctaefectosdesc|||"
      Text            =   "Text1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   13
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "Text2"
      Top             =   3960
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   120
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Cta. gastos|T|S|||ctabancaria|ctagastostarj|||"
      Text            =   "Text1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   12
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   3960
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "idnorma34|T|S|||ctabancaria|idnorma34|||"
      Text            =   "Text1"
      Top             =   1800
      Width           =   1845
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   10
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "Text2"
      Top             =   3240
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   14
      Tag             =   "Cta. gastos|T|S|||ctabancaria|ctaingreso|||"
      Text            =   "Text1"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   8
      Tag             =   "Sufijo OEM|T|S|||ctabancaria|sufijoem|||"
      Text            =   "Text1"
      Top             =   1800
      Width           =   645
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "Text2"
      Top             =   6720
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   5160
      MaxLength       =   4
      TabIndex        =   26
      Tag             =   "Centro Coste|T|S|||ctabancaria|codccost|||"
      Text            =   "Text1"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   120
      MaxLength       =   15
      TabIndex        =   6
      Tag             =   "Cedante|T|S|||ctabancaria|idCedente|||"
      Text            =   "Text1"
      Top             =   1800
      Width           =   1845
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Cuenta bancaria"
      Height          =   615
      Left            =   5520
      TabIndex        =   41
      Top             =   720
      Width           =   4335
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   20
         Left            =   120
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "Entidad|T|S|||ctabancaria|iban|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Oficina|N|N|||ctabancaria|oficina|0000||"
         Text            =   "Text1"
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "Entidad|T|N|||ctabancaria|entidad|0000||"
         Text            =   "Text1"
         Top             =   240
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   2460
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "Digito control|T|S|||ctabancaria|control|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Cuenta banco|T|N|||ctabancaria|ctabanco|0000000000||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
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
         Index           =   24
         Left            =   180
         TabIndex        =   72
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad"
         Height          =   195
         Index           =   2
         Left            =   1500
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Oficina"
         Height          =   195
         Index           =   4
         Left            =   2340
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "D.C."
         Height          =   195
         Index           =   6
         Left            =   3060
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Cta banco"
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   120
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "Cta. gastos|T|S|||ctabancaria|ctagastos|||"
      Text            =   "Text1"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   3240
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Tag             =   "Cta. contable|T|N|||ctabancaria|codmacta||S|"
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   960
      Width           =   3795
   End
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9075
      TabIndex        =   30
      Top             =   7875
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   5280
      MaxLength       =   40
      TabIndex        =   9
      Tag             =   "Descripcion|T|S|||ctabancaria|descripcion|||"
      Text            =   "Text1"
      Top             =   1800
      Width           =   4965
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   32
      Top             =   7920
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9090
      TabIndex        =   31
      Top             =   7875
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7770
      TabIndex        =   29
      Top             =   7875
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   8280
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7800
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Docum. Pagaré"
      Height          =   195
      Index           =   29
      Left            =   2880
      TabIndex        =   77
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Docum. talon"
      Height          =   255
      Index           =   28
      Left            =   120
      TabIndex        =   76
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "NºContrato"
      Height          =   255
      Index           =   27
      Left            =   6360
      TabIndex        =   75
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "SEPA 19.14 Empresa deudora graba NIF"
      Height          =   255
      Index           =   26
      Left            =   2280
      TabIndex        =   74
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   5280
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Sufijo N34 SEPA"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   73
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Gastos transferencia descontados importe"
      Height          =   255
      Index           =   23
      Left            =   600
      TabIndex        =   71
      Top             =   6840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Gastos remesa descontados importe"
      Height          =   255
      Index           =   20
      Left            =   600
      TabIndex        =   68
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   13
      Left            =   6600
      Picture         =   "frmBanco.frx":000C
      ToolTipText     =   "Cta efectos descontados"
      Top             =   3720
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta efec. desconta."
      Height          =   195
      Index           =   13
      Left            =   5280
      TabIndex        =   56
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   12
      Left            =   1440
      Picture         =   "frmBanco.frx":685E
      ToolTipText     =   "Cuenta tarjeta"
      Top             =   3720
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta gastos tarjeta"
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   54
      Top             =   3720
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Id. Norma34"
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   52
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta ingresos"
      Height          =   195
      Index           =   10
      Left            =   5280
      TabIndex        =   51
      Top             =   3000
      Width           =   870
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   10
      Left            =   6240
      Picture         =   "frmBanco.frx":D0B0
      ToolTipText     =   "Cuenta ingresos"
      Top             =   3000
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Sufijo EM"
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   49
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Centro coste"
      Height          =   195
      Index           =   8
      Left            =   5160
      TabIndex        =   48
      Top             =   6465
      Width           =   900
   End
   Begin VB.Image imgCC 
      Height          =   240
      Left            =   6240
      Picture         =   "frmBanco.frx":13902
      Top             =   6420
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Id. Cedente"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   5
      Left            =   960
      Picture         =   "frmBanco.frx":1A154
      ToolTipText     =   "Cuenta gastos"
      Top             =   3000
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta gastos"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   40
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   4
      Left            =   1440
      Picture         =   "frmBanco.frx":209A6
      ToolTipText     =   "Cuenta contable"
      Top             =   660
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta contable"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   38
      Top             =   705
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   34
      Top             =   1560
      Width           =   975
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private DevfrmCCtas As String



Private Sub chkBanco_KeyPress(Index As Integer, KeyAscii As Integer)
 KeyPressGral KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                'MsgBox "Registro insertado.", vbInformation
                PonerModo 0
                lblIndicador.Caption = ""
            End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    lblIndicador.Caption = ""
                    If SituarData1 Then
                        PonerModo 2
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    'Modificar
    lblIndicador.Caption = ""
    TerminaBloquear
    PonerModo 2
    PonerCampos
End Select

End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim SQL As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " codmacta = " & Text1(4).Text & ""
            Data1.Recordset.Find SQL
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    SugerirCodigoSiguiente
    '###A mano
    Text1(4).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(4).SetFocus
        Text1(4).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
    Case 1
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 3
        Data1.Recordset.MoveLast
End Select
PonerCampos
lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(4).Locked = True
    Text1(4).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    Text1(1).SetFocus
End Sub

Private Sub BotonEliminar()

'
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    'Comprobamos si se puede eliminar
    I = 0
    If Not SePuedeEliminar Then I = 1
     
    Set miRsAux = Nothing
    If I = 1 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '### a mano
    Cad = "Seguro que desea eliminar de la BD el registro:"
    Cad = Cad & vbCrLf & "Cta banco: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Decripcion: " & Me.Text2(4).Text
    I = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            Else
                Data1.Recordset.MoveFirst
                NumRegElim = NumRegElim - 1
                If NumRegElim > 1 Then
                    For I = 1 To NumRegElim - 1
                        Data1.Recordset.MoveNext
                    Next I
                End If
                PonerCampos
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub




Private Sub cmdRegresar_Click()

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If



    
    

    RaiseEvent DatoSeleccionado(CStr(Text1(4).Text & "|" & Text2(4).Text & "|"))
    Unload Me
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub




Private Sub Form_Load()
Dim I As Integer


      ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    Me.Icon = frmPpal.Icon

    LimpiarCampos

    'Como son cuentas, como mucho seran
    For I = 4 To 5
        Text1(I).MaxLength = vEmpresa.DigitosUltimoNivel
    Next I
    
    '## A mano
    NombreTabla = "ctabancaria"
    Ordenacion = " ORDER BY codmacta"
        
    PonerOpcionesMenu
    
    'Para todos
'    Data1.UserName = vUsu.Login
'    Me.Data1.password = vUsu.Passwd
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        '### A mano
        'Text1(0).BackColor = vbYellow
    End If
    FrameAnalitica.Visible = Not vParam.autocoste
    If Not vParam.autocoste Then Me.Text1(8).TabIndex = 100
    'If vParam.autocoste Then Text1(8).TabIndex = 9
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano

    'Check1.Value = 0
    For kCampo = 0 To 2
        Me.chkBanco(kCampo).Value = 0
    Next
    kCampo = 0
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim AUX As String
    If CadenaDevuelta <> "" Then
    
        If Me.Tag = "" Then
                    
                HaDevueltoDatos = True
                Screen.MousePointer = vbHourglass
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                CadB = ""
                AUX = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
                CadB = AUX
                '   Como la clave principal es unica, con poner el sql apuntando
                '   al valor devuelto sobre la clave ppal es suficiente
                'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                'If CadB <> "" Then CadB = CadB & " AND "
                'CadB = CadB & Aux
                'Se muestran en el mismo form
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda
                Screen.MousePointer = vbDefault
                
        Else
            'Es el busqueda de los CC
            Text1(8).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(2).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub imgCC_Click()
    'Lanzaremos el vista previa
    Screen.MousePointer = vbHourglass
        'Cod Diag.|idDiag|N|10·
        DevfrmCCtas = "C.C.|codccost|T|25·Descripcion C.C.|nomccost|T|65·"
        
        Me.Tag = "CC"
        Set frmB = New frmBuscaGrid
        frmB.vCampos = DevfrmCCtas
        frmB.vTabla = "cabccost"
        frmB.vSQL = ""
        DevfrmCCtas = ""
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Centros de Coste"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing

End Sub

Private Sub imgCuentas_Click(Index As Integer)
 Screen.MousePointer = vbHourglass
 Set frmCCtas = New frmColCtas
 DevfrmCCtas = ""
 frmCCtas.DatosADevolverBusqueda = "0"
 frmCCtas.Show vbModal
 Set frmCCtas = Nothing
 If DevfrmCCtas <> "" Then
        Text1(Index).Text = RecuperaValor(DevfrmCCtas, 1)
        Text2(Index).Text = RecuperaValor(DevfrmCCtas, 2)
End If
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
        Else
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub


'Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        KeyCode = 0
'        SendKeys "{TAB}"
'    End If
'End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyPressGral KeyAscii
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim Valor As Currency
    Dim SQL As String
    Dim mTag As CTag
    Dim I As Integer
    ''Quitamos blancos por los lados
    If Index <> 11 Then Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    
    If Modo = 1 Then Exit Sub
    'Si queremos hacer algo ..
    Select Case Index
        Case 0, 2, 3, 6
            If Modo = 3 Or Modo = 4 Then
                If Text1(Index).Text = "" Then Exit Sub
                Set mTag = New CTag
                If mTag.Cargar(Text1(Index)) Then
                    If mTag.Cargado Then
                        If mTag.Comprobar(Text1(Index)) Then
                            FormateaCampo Text1(Index), mTag 'Formateamos el campo si tiene valor
                        Else
                            Text1(Index).Text = ""
                            Ponerfoco Text1(Index)
                        End If
                    End If
                End If
                Set mTag = Nothing
            End If
             
             
            SQL = Text1(2).Text & Text1(3).Text & Text1(6).Text & Text1(0).Text
                    
            If Len(SQL) = 20 Then
                'OK. Calculamos el IBAN
                
                
                If Text1(20).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", SQL, SQL) Then Text1(20).Text = "ES" & SQL
                Else
                    DevfrmCCtas = CStr(Mid(Text1(20).Text, 1, 2))
                    If DevuelveIBAN2(DevfrmCCtas, SQL, SQL) Then
                        If Mid(Text1(20).Text, 3) <> SQL Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & DevfrmCCtas & SQL & "]", vbExclamation
                            'Text1(49).Text = "ES" & SQL
                        End If
                    End If
                    DevfrmCCtas = ""
                End If
            End If
                
             
             
        Case 20
            'IBAN
            If Text1(Index).Text <> "" Then
                If Not IBAN_Correcto(Me.Text1(Index).Text) Then Text1(Index).Text = ""
            End If
        Case 4, 5, 10, 12, 13
            
            If Modo >= 2 Or Modo <= 4 Then
                If Text1(Index).Text = "" Then
                     Text2(Index).Text = SQL
                     Exit Sub
                End If
                
                DevfrmCCtas = Text1(Index).Text
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                    Text1(Index).Text = DevfrmCCtas
                    Text2(Index).Text = SQL
                Else
                    MsgBox SQL, vbExclamation
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                    Ponerfoco Text1(Index)
                End If
                DevfrmCCtas = ""
                
            End If
        Case 8
            If Text1(8).Text = "" Then
                Text2(2).Text = ""
                Exit Sub
            End If
            DevfrmCCtas = DevuelveDesdeBD("nomccost", "cabccost", "codccost", Text1(8).Text, "T")
            If DevfrmCCtas = "" Then
                MsgBox "CC no encontrado: " & Text1(8).Text, vbExclamation
                Text1(8).Text = ""
                Exit Sub
            Else
                Text1(8).Text = UCase(Text1(8).Text)
            End If
            Text2(2).Text = DevfrmCCtas
            
        Case 14, 17, 21, 22
            'Dias
            Text1(Index).Text = Trim(Text1(Index).Text)
            If Text1(Index).Text = "" Then Exit Sub
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Campo numérico: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
            Else
                Text1(Index).Text = Abs(Val(Text1(Index).Text))
            End If
        Case 15, 16, 18, 19
            'Importe
            Text1(Index).Text = Trim(Text1(Index).Text)
            If Text1(Index).Text = "" Then Exit Sub
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "importe debe ser numérico", vbExclamation
                Text1(Index).Text = ""
                Ponerfoco Text1(Index)
            Else
                If InStr(1, Text1(Index).Text, ",") > 0 Then
                    Valor = ImporteFormateado(Text1(Index).Text)
                Else
                    Valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                End If
                Text1(Index).Text = Format(Valor, FormatoImporte)
            End If
                
            
        '....
    End Select
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me)

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(4), 20, "Cuenta")
        Cad = Cad & ParaGrid(Text1(1), 40, "Denominacion")
        Cad = Cad & ParaGrid(Text1(2), 10, "Entidad")
        Cad = Cad & ParaGrid(Text1(3), 10, "Oficina")
        Cad = Cad & ParaGrid(Text1(0), 20, "Cta. banco")
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Me.Tag = ""
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Cuenta bancos propios"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
            End If
        End If
End Sub



Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
    Screen.MousePointer = vbDefault
    Exit Sub

    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
End If


Screen.MousePointer = vbDefault
Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim I As Integer
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    PonerCtasIVA
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Dim B As Boolean
    Dim Obj
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            'Text1(i).BackColor = vbWhite
            Text1(0).BackColor = &H80000018
        Next I
        'chkVistaPrevia.Visible = False
    End If
    Modo = Kmodo
    'chkVistaPrevia.Visible = (Modo = 1)
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    'Modificar
    Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
    mnEliminar.Enabled = B
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = B Or Modo = 1
    cmdCancelar.Visible = B Or Modo = 1
    mnOpciones.Enabled = Not B
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 2
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = (Modo = 2) Or Modo = 0
    For I = 0 To 26
        
            Text1(I).Locked = B
            If B Then
                Text1(I).BackColor = &H80000018
            ElseIf Modo <> 1 Then
                Text1(I).BackColor = vbWhite
            End If
        
    Next I
    Me.chkBanco(0).Enabled = Not B
    Me.chkBanco(1).Enabled = Not B
    Me.chkBanco(2).Enabled = Not B
    
    For Each Obj In imgCuentas
        Obj.Visible = Not B
    Next
    Me.imgCC.Visible = Not B
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    'Si tiene contabilidad analitica EXITGIMOS EL CC
    If vParam.autocoste Then
        If Text1(8).Text = "" Then
            MsgBox "Centro de coste requerido", vbExclamation
            Exit Function
        End If
    End If
    
    If Text1(2).Text <> "" Then
        If Val(Text1(2).Text) <> 0 Then
            If CodigoDeControl(Text1(2).Text & Text1(3).Text & Text1(0).Text) <> Text1(6).Text Then
                If MsgBox("Codigo control incorrecto (" & CodigoDeControl(Text1(2).Text & Text1(3).Text & Text1(0).Text) & ") ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
                
           
     
                
                
        End If
    End If
    
    
    'Si el idNorma34 son espacios en blanco entonces pong "", para que en la BD vaya un NULL
    If Trim(Text1(11).Text) = "" Then Text1(11).Text = ""
    
    'Comprobamos  si existe
    If Modo = 3 Then
       ' If DevuelveDesdeBD("codigiva", "tiposiva", "codigiva", Text1(0).Text, "N") <> "" Then
       '     B = False
       '     MsgBox "Ya existe el codigo de IVA: " & Text1(0).Text, vbExclamation
       ' Else
       '     B = True
       ' End If
    End If
    DatosOk = B
End Function


'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
'
'    Dim SQL As String
'    Dim RS As ADODB.Recordset
'
'    SQL = "Select Max(codigiva) from " & NombreTabla
'    Text1(0).Text = 1
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, , , adCmdText
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            Text1(0).Text = RS.Fields(0) + 1
'        End If
'    End If
'    RS.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Index
Case 1
    BotonBuscar
Case 2
    BotonVerTodos
Case 6
    BotonAnyadir
Case 7
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
Case 8
    BotonEliminar
Case 12
    mnSalir_Click
Case 14 To 17
    Desplazamiento (Button.Index - 14)
Case 11
    
    If ListadoCtaBanco Then
        'Imprimimimos
        With frmImprimir
            .Opcion = 26
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .NumeroParametros = 0
            .SoloImprimir = False
            .Show vbModal
        End With
    End If
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub


Private Sub PonerCtasIVA()
On Error GoTo EPonerCtasIVA

    Text1_LostFocus 4
    Text1_LostFocus 5
    Text1_LostFocus 8
    Text1_LostFocus 10
    Text1_LostFocus 12
    Text1_LostFocus 13
Exit Sub
EPonerCtasIVA:
    MuestraError Err.Number, "Poniendo valores ctas.", Err.Description
End Sub



Private Sub Ponerfoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SePuedeEliminar() As Boolean
Dim B As Boolean
Dim Cad As String

    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    SePuedeEliminar = False
    
    'Veamos cobros asociados
    Cad = "Select count(*) from scobro where (cuentaba = '" & Data1.Recordset.Fields(0) & "' or ctabanc2 = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Banco con cobros pendientes", vbExclamation
        Exit Function
    End If
    
    
    
    Cad = "Select count(*) from spagop where (ctabanc1 = '" & Data1.Recordset.Fields(0) & "' or ctabanc2 = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Banco con pagos pendientes", vbExclamation
        Exit Function
    End If
    
    'Remesas
    Cad = "Select count(*) from remesas where (codmacta = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Remesas asociadas.", vbExclamation
        Exit Function
    End If
    
    
    Cad = "Select count(*) from sgastfij where (ctaprevista = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Gasto fijo asociado.", vbExclamation
        Exit Function
    End If
    
    
    
    Cad = "Select count(*) from stransfer where (codmacta= '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Transferencia pagos asociada.", vbExclamation
        Exit Function
    End If
    
        
    
    
    Cad = "Select count(*) from stransfercob where (codmacta= '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Transferencia abono asociada.", vbExclamation
        Exit Function
    End If
    
    
    'cOMPROBAMOS ai tiene moovimientos en
    'la NORMA 43
    Cad = "Select count(*) from norma43 where (codmacta= '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Asociada a norma 43 en la contabilidad.", vbExclamation
        Exit Function
    End If
    
    SePuedeEliminar = True
    Screen.MousePointer = vbDefault
End Function

