VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCobros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14025
   Icon            =   "frmCobros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12840
      TabIndex        =   43
      Top             =   8595
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12840
      TabIndex        =   59
      Top             =   8595
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11640
      TabIndex        =   42
      Top             =   8595
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4560
      Top             =   8640
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
      TabIndex        =   62
      Top             =   0
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
            Object.ToolTipText     =   "Ver impagados"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Efectuar cobro "
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Dividir vencimiento"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8040
         TabIndex        =   63
         Top             =   120
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   69
      Top             =   1320
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmCobros.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(31)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(30)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(29)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(28)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgCuentas(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgCuentas(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgFecha(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "imgCuentas(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(10)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(12)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgDepart"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(11)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(18)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(19)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(14)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "imgCuentas(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "imgFecha(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(21)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "imgAgente"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(27)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "imgFecha(7)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(33)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(34)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(35)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "FrameEstaEnCaja"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(31)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(30)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(29)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(28)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text2(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text2(1)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(5)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(6)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Frame2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text2(2)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(9)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text2(3)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(10)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Check1(1)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text1(33)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text2(4)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "FrameRemesa"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text1(32)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Text1(38)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text1(39)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "frameContene"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Text1(40)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Text2(5)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Text1(34)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Text1(42)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtPendiente"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "FrameSeguro"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Text1(48)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Text1(49)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Text1(12)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Text1(11)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).ControlCount=   65
      TabCaption(1)   =   "Textos CSB / Referencia"
      TabPicture(1)   =   "frmCobros.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text3(0)"
      Tab(1).Control(1)=   "Text3(1)"
      Tab(1).Control(2)=   "Text1(15)"
      Tab(1).Control(3)=   "Text1(14)"
      Tab(1).Control(4)=   "Text1(17)"
      Tab(1).Control(5)=   "Text1(16)"
      Tab(1).Control(6)=   "Text1(19)"
      Tab(1).Control(7)=   "Text1(18)"
      Tab(1).Control(8)=   "Text1(21)"
      Tab(1).Control(9)=   "Text1(20)"
      Tab(1).Control(10)=   "Text1(23)"
      Tab(1).Control(11)=   "Text1(22)"
      Tab(1).Control(12)=   "Text1(25)"
      Tab(1).Control(13)=   "Text1(24)"
      Tab(1).Control(14)=   "Text1(27)"
      Tab(1).Control(15)=   "Text1(26)"
      Tab(1).Control(16)=   "Line5"
      Tab(1).Control(17)=   "Line4"
      Tab(1).Control(18)=   "Line3"
      Tab(1).Control(19)=   "Label2"
      Tab(1).Control(20)=   "Label3(0)"
      Tab(1).Control(21)=   "Line1"
      Tab(1).Control(22)=   "Label3(1)"
      Tab(1).Control(23)=   "Label3(2)"
      Tab(1).Control(24)=   "Line2"
      Tab(1).ControlCount=   25
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   11
         Left            =   360
         MaxLength       =   80
         TabIndex        =   36
         Tag             =   "CSB|T|S|||scobro|text33csb|||"
         Text            =   "WWW4567890WWW4567890WWW4567890WWW456789WWWW4567890WWW4567890WWW4567890WWW456789J"
         Top             =   5520
         Width           =   7965
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   360
         MaxLength       =   60
         TabIndex        =   37
         Tag             =   "T|T|S|||scobro|text41csb|||"
         Top             =   6360
         Width           =   6765
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   49
         Left            =   360
         MaxLength       =   4
         TabIndex        =   23
         Tag             =   "Iban|T|S|||scobro|iban|||"
         Text            =   "ES99"
         Top             =   3990
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   48
         Left            =   12000
         TabIndex        =   32
         Tag             =   "Fecha ult ejecucion|F|S|||scobro|fecejecutiva|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   1455
      End
      Begin VB.Frame FrameSeguro 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   735
         Left            =   240
         TabIndex        =   112
         Top             =   4440
         Width           =   8295
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   47
            Left            =   3480
            TabIndex        =   35
            Tag             =   "Aviso siniestro|F|S|||scobro|fecsiniestro|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   330
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   46
            Left            =   1800
            TabIndex        =   34
            Tag             =   "Aviso prorroga|F|S|||scobro|fecprorroga|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   330
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   45
            Left            =   120
            TabIndex        =   33
            Tag             =   "Fecha Aviso falta pago|F|S|||scobro|feccomunica|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblAsegurado 
            Caption         =   "Asegurado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   5400
            TabIndex        =   116
            Top             =   120
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "F. aviso siniestro"
            Height          =   195
            Index           =   26
            Left            =   3480
            TabIndex        =   115
            Top             =   120
            Width           =   1395
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   6
            Left            =   4710
            Top             =   405
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3030
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Prorroga"
            Height          =   195
            Index           =   25
            Left            =   1800
            TabIndex        =   114
            Top             =   120
            Width           =   1155
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   1350
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F. aviso falta pago"
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   113
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPendiente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEF7E4&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9360
         TabIndex        =   110
         Text            =   "Text4"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   42
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   30
         Tag             =   "Ref|T|S|||scobro|reftalonpag|||"
         Text            =   "Text1"
         Top             =   4005
         Width           =   2085
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   34
         Left            =   360
         TabIndex        =   7
         Tag             =   "Agente|N|N|0||scobro|agente|||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "Text2"
         Top             =   1500
         Width           =   3435
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   40
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Transferencia|N|S|0||scobro|transfer|0000000000||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   1125
      End
      Begin VB.Frame frameContene 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7440
         TabIndex        =   103
         Top             =   3000
         Width           =   6135
         Begin VB.CheckBox Check1 
            Caption         =   "Documento recibido"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   19
            Tag             =   "Recibido|N|S|||scobro|recedocu|||"
            Top             =   120
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Devuelto"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   20
            Tag             =   "Devuelto|N|S|||scobro|Devuelto|||"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Situacion jurídica"
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   22
            Tag             =   "s|N|S|||scobro|situacionjuri|||"
            Top             =   120
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "NO remesar"
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   21
            Tag             =   "s|N|S|||scobro|noremesar|||"
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1275
         Index           =   39
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   38
         Tag             =   "obs|T|S|||scobro|obs|||"
         Text            =   "Text1"
         Top             =   5520
         Width           =   4965
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   38
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   12
         Tag             =   "Gastos|N|S|||scobro|gastos|#,##0.00||"
         Text            =   "1.999.999.00"
         Top             =   2340
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   32
         Left            =   4680
         TabIndex        =   28
         Tag             =   "Ultima reclamacion|F|S|||scobro|ultimareclamacion|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   3990
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -74520
         TabIndex        =   99
         Text            =   "Text3"
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -69840
         TabIndex        =   98
         Text            =   "Text3"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Frame FrameRemesa 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   735
         Left            =   360
         TabIndex        =   93
         Top             =   2880
         Width           =   6615
         Begin VB.ComboBox cboTipoRem 
            Height          =   315
            ItemData        =   "frmCobros.frx":0044
            Left            =   960
            List            =   "frmCobros.frx":0051
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "Remesa|N|S|0||scobro|tiporem|||"
            Top             =   290
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   37
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   18
            Tag             =   "Situacion|T|S|0||scobro|siturem|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   36
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "Año remesa|N|S|0||scobro|anyorem|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   885
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   35
            Left            =   3240
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Remesa|N|S|0||scobro|codrem|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Situacion"
            Height          =   195
            Index           =   17
            Left            =   5520
            TabIndex        =   97
            Top             =   90
            Width           =   780
         End
         Begin VB.Label Label4 
            Caption         =   "REMESA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   96
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Año"
            Height          =   195
            Index           =   16
            Left            =   4440
            TabIndex        =   95
            Top             =   90
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "Numero"
            Height          =   195
            Index           =   15
            Left            =   3240
            TabIndex        =   94
            Top             =   90
            Width           =   660
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "Text2"
         Top             =   720
         Width           =   3195
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   33
         Left            =   4680
         TabIndex        =   5
         Tag             =   "departamento|N|S|||scobro|departamento|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Variable CONTDOCU"
         Height          =   255
         Index           =   1
         Left            =   11400
         TabIndex        =   44
         Tag             =   "Doc. Contabilizado|N|S|||scobro|contdocu|||"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   46
         Tag             =   "T|T|S|||scobro|text43csb|||"
         Top             =   2400
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   45
         Tag             =   "T|T|S|||scobro|text42csb|||"
         Top             =   2400
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   48
         Tag             =   "T|T|S|||scobro|text52csb|||"
         Top             =   2760
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   16
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "T|T|S|||scobro|text51csb|||"
         Top             =   2760
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   19
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   50
         Tag             =   "T|T|S|||scobro|text61csb|||"
         Top             =   3120
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   49
         Tag             =   "T|T|S|||scobro|text53csb|||"
         Top             =   3120
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   52
         Tag             =   "T|T|S|||scobro|text63csb|||"
         Top             =   3480
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   20
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   51
         Tag             =   "T|T|S|||scobro|text62csb|||"
         Top             =   3480
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   23
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   54
         Tag             =   "T|T|S|||scobro|text72csb|||"
         Top             =   3840
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   53
         Tag             =   "T|T|S|||scobro|text71csb|||"
         Top             =   3840
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   25
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   56
         Tag             =   "T|T|S|||scobro|text81csb|||"
         Top             =   4200
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   24
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   55
         Tag             =   "T|T|S|||scobro|text73csb|||"
         Top             =   4200
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   -70200
         MaxLength       =   40
         TabIndex        =   58
         Tag             =   "T|T|S|||scobro|text83csb|||"
         Top             =   4560
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   26
         Left            =   -74520
         MaxLength       =   40
         TabIndex        =   57
         Tag             =   "T|T|S|||scobro|text82csb|||"
         Top             =   4560
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   9360
         TabIndex        =   9
         Tag             =   "Cta real pago|T|S|||scobro|ctabanc2|||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text2"
         Top             =   1500
         Width           =   3075
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   4680
         TabIndex        =   8
         Tag             =   "Cta prevista|T|N|||scobro|ctabanc1|||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text2"
         Top             =   1500
         Width           =   3195
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   4680
         TabIndex        =   72
         Top             =   2010
         Width           =   4575
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Importe|N|S|||scobro|impcobro|#,##0.00||"
            Text            =   "1.999.999.00"
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   7
            Left            =   1440
            TabIndex        =   13
            Tag             =   "Fecha ult. pago|F|S|||scobro|fecultco|dd/mm/yyyy||"
            Text            =   "99/99/9999"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ultimo pago"
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   120
            Top             =   120
            Width           =   1155
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   2040
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Importe"
            Height          =   195
            Index           =   8
            Left            =   2880
            TabIndex        =   74
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   7
            Left            =   1440
            TabIndex        =   73
            Top             =   120
            Width           =   555
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "Importe|N|N|||scobro|impvenci|#,##0.00||"
         Text            =   "1.999.999.00"
         Top             =   2340
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Tag             =   "Fecha vencimiento|F|N|||scobro|fecvenci|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   2340
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   9360
         TabIndex        =   6
         Tag             =   "Forma Pago|N|N|0||scobro|codforpa|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   720
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   4
         Tag             =   "Cta. cliente|T|N|||scobro|codmacta|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   720
         Width           =   3195
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   28
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   24
         Tag             =   "Entidad|N|S|0||scobro|codbanco|0000||"
         Text            =   "9999"
         Top             =   3990
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   29
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   25
         Tag             =   "Sucursal|N|S|0||scobro|codsucur|0000||"
         Text            =   "9999"
         Top             =   3990
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   30
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   26
         Tag             =   "D.C.|T|S|0||scobro|digcontr|||"
         Text            =   "99"
         Top             =   3990
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   31
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Cuenta|T|S|||scobro|cuentaba|0000000000||"
         Text            =   "9999999999"
         Top             =   3990
         Width           =   1125
      End
      Begin VB.Frame FrameEstaEnCaja 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Left            =   9960
         TabIndex        =   105
         Top             =   3720
         Width           =   1455
         Begin VB.CheckBox Check1 
            Caption         =   "Esta en caja"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   31
            Tag             =   "s|N|S|||scobro|estacaja|||"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Linea2 SEPA"
         Height          =   195
         Index           =   35
         Left            =   360
         TabIndex        =   123
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Linea1 SEPA"
         Height          =   195
         Index           =   34
         Left            =   360
         TabIndex        =   122
         Top             =   5280
         Width           =   1335
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
         Index           =   33
         Left            =   360
         TabIndex        =   121
         Top             =   3720
         Width           =   780
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   13200
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Reca. ejecutiva"
         Height          =   195
         Index           =   27
         Left            =   12000
         TabIndex        =   117
         Top             =   3795
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   111
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image imgAgente 
         Height          =   255
         Left            =   960
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Num. ref. talón/pagare"
         Height          =   195
         Index           =   21
         Left            =   7560
         TabIndex        =   109
         Top             =   3795
         Width           =   1860
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   5520
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   10800
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   107
         Top             =   1290
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Transferencia"
         Height          =   195
         Index           =   19
         Left            =   6120
         TabIndex        =   104
         Top             =   3780
         Width           =   1020
      End
      Begin VB.Label Label5 
         Caption         =   "Obs:"
         Height          =   255
         Left            =   8520
         TabIndex        =   102
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Gastos"
         Height          =   195
         Index           =   18
         Left            =   3120
         TabIndex        =   101
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Ult. reclama."
         Height          =   195
         Index           =   11
         Left            =   4680
         TabIndex        =   100
         Top             =   3780
         Width           =   915
      End
      Begin VB.Image imgDepart 
         Height          =   240
         Left            =   5760
         ToolTipText     =   "Buscar departamento"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Index           =   12
         Left            =   4680
         TabIndex        =   92
         Top             =   480
         Width           =   1005
      End
      Begin VB.Line Line5 
         X1              =   -74640
         X2              =   -74640
         Y1              =   960
         Y2              =   5040
      End
      Begin VB.Line Line4 
         X1              =   -65880
         X2              =   -65880
         Y1              =   960
         Y2              =   5040
      End
      Begin VB.Line Line3 
         X1              =   -74640
         X2              =   -65880
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   -74640
         TabIndex        =   90
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CARGO POR DOMICILIACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -69960
         TabIndex        =   89
         Top             =   600
         Width           =   4095
      End
      Begin VB.Line Line1 
         X1              =   -74640
         X2              =   -65880
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ORDENANTE"
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
         Left            =   -74520
         TabIndex        =   88
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TITULAR"
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
         Left            =   -70200
         TabIndex        =   87
         Top             =   1320
         Width           =   810
      End
      Begin VB.Line Line2 
         X1              =   -74640
         X2              =   -65880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. real de cobro"
         Height          =   195
         Index           =   10
         Left            =   9360
         TabIndex        =   86
         Top             =   1260
         Width           =   1260
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   6120
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. prevista cobro"
         Height          =   195
         Index           =   9
         Left            =   4680
         TabIndex        =   85
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1200
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   84
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Vto."
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   83
         Top             =   2100
         Width           =   795
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   10560
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         Height          =   195
         Index           =   0
         Left            =   9360
         TabIndex        =   82
         Top             =   480
         Width           =   1065
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1200
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. cliente"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   81
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad"
         Height          =   195
         Index           =   28
         Left            =   1200
         TabIndex        =   80
         Top             =   3780
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal"
         Height          =   195
         Index           =   29
         Left            =   1800
         TabIndex        =   79
         Top             =   3780
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "D.C."
         Height          =   195
         Index           =   30
         Left            =   2400
         TabIndex        =   78
         Top             =   3780
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   31
         Left            =   2880
         TabIndex        =   77
         Top             =   3780
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   60
      Top             =   8520
      Width           =   4095
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Frame FrameClaves 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   615
      Left            =   120
      TabIndex        =   64
      Top             =   600
      Width           =   13815
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   43
         Left            =   9360
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Referencia1|T|S|||scobro|referencia1|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   44
         Left            =   11640
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Referencia2|T|S|||scobro|referencia2|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   41
         Left            =   7320
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Referencia|T|S|0||scobro|referencia|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   13
         Left            =   120
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "Serie|T|N|||scobro|numserie||S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Nº Factura|N|N|||scobro|codfaccl|000000|S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Nº Vencimiento|N|N|0||scobro|numorden||S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||scobro|fecfaccl|dd/mm/yyyy|S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia (I)"
         Height          =   195
         Index           =   22
         Left            =   9360
         TabIndex        =   119
         Top             =   30
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia (II)"
         Height          =   195
         Index           =   23
         Left            =   11640
         TabIndex        =   118
         Top             =   30
         Width           =   1020
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Left            =   600
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia"
         Height          =   195
         Index           =   20
         Left            =   7320
         TabIndex        =   108
         Top             =   30
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   68
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Nº  Factura"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   67
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Vencimiento"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   66
         Top             =   30
         Width           =   1140
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3240
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Factura"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   65
         Top             =   0
         Width           =   915
      End
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
Attribute VB_Name = "frmCobros"
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
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmD As frmDepartamentos
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmF As frmFormaPago
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmS As frmSerie
Attribute frmS.VB_VarHelpID = -1
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

'NUEVO: DICIEMBRE 2005. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String


Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
    
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
                Cad = DameClavesADODCForm(Me, Me.Data1)
                
                If ModificaDesdeFormularioClaves(Me, Cad) Then
                    
                    'TerminaBloquear
                    DesBloqueaRegistroForm Me.Text1(0)
                    lblIndicador.Caption = ""
                    If SituarData Then
                    
                        Text1_LostFocus 0
                        Cad = Text2(1).Tag 'para que no pierda el valor
                        PonerModo 2
                        Text2(1).Tag = Cad
                        Cad = ""
                        PonPendiente
                        '-- Esto permanece para saber donde estamos
                        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
                    Else
                        LimpiarCampos
                        'PonerModo 0
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
    'TerminaBloquear
    DesBloqueaRegistroForm Me.Text1(0)
    PonerModo 2
    PonerCampos
End Select

End Sub



Private Function SituarData() As Boolean
    Dim Posicion As Long
    Dim SQL As String
    On Error GoTo ESituarData1
        SituarData = False
                    
        With Data1
            'Vemos poscion
            Posicion = .Recordset.AbsolutePosition - 1
            'Actualizamos el recordset
            .Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            .Recordset.MoveFirst
            
            If .Recordset.RecordCount <= Posicion Then
                'Era el utlimo
                .Recordset.MoveLast
            Else
                If Posicion > 0 Then .Recordset.Move Posicion
            End If
            SituarData = True
'            While Not .Recordset.EOF
'                If .Recordset!NUmSerie = Text1(13).Text Then
'                    If .Recordset!codfaccl = Text1(1).Text Then
'                        If Format(.Recordset!fecfaccl, "dd/mm/yyyy") = Text1(2).Text Then
'                            If CStr(.Recordset!numorden) = Text1(3).Text Then
'                                SituarData = True
'                                Exit Function
'                            End If
'                        End If
'                    End If
'                End If
'                .Recordset.MoveNext
'            Wend
        End With
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    Check1(1).Value = 0
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    
    '###A mano
    Text1(13).SetFocus
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
        Text1(13).SetFocus
        Text1(13).BackColor = vbYellow
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
        CadenaConsulta = "Select * from " & NombreTabla
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
Dim N As Byte

    N = SePuedeEliminar2()
    If N = 0 Then Exit Sub


    If Not BloqueaRegistroForm(Me) Then Exit Sub
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    
    'Si se puede modificar entonces habilito todooos los campos
    PonerModo 4
    If N < 3 Then
        'Se puede modifcar la CC
        Dim T As TextBox
        For Each T In Text1
            If T.Index < 28 Or T.Index > 31 Then
                T.Locked = True
                T.BackColor = &H80000018
            End If
        Next T
        'Tabbien dejamos modificar el IBAN
        Text1(49).Locked = False
        Text1(49).BackColor = vbWhite
        'Pongo visible false los img
         For N = 0 To 6
            If N < 4 Then imgCuentas(N).Visible = False
            Me.imgFecha(N).Visible = False
         Next N
        
        
        'Si es una remesa de talon/pagare tb dejare modificar el numero de talon pagare
        If Val(DBLet(Data1.Recordset!Tiporem)) > 1 Then
            Text1(42).Locked = False
            Text1(42).BackColor = vbWhite
        End If
            
        Ponerfoco Text1(28)
    Else
        Ponerfoco Text1(6)
    End If
    
    
    'Si no tienen permisos NO permito modificar
    If vParam.TieneOperacionesAseguradas Then
        If vUsu.Nivel >= 1 Then FrameSeguro.Enabled = False
    End If
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
'    Text1(0).Locked = True
'    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'Comprobamos si se puede eliminar
    I = SePuedeEliminar2
    If I < 3 Then Exit Sub
    '### a mano
    Cad = "Seguro que desea eliminar de la BD el registro actual:"
    Cad = Cad & vbCrLf & Data1.Recordset.Fields(0) & "  " & Data1.Recordset.Fields(1) & " "
    Cad = Cad & Data1.Recordset.Fields(2) & "  " & Data1.Recordset.Fields(3)
    I = MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton2)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        
        'para sefectdev
        Cad = "DELETE FROM sefecdev WHERE numserie = '" & Data1.Recordset!NUmSerie & "' AND codfaccl = " & Data1.Recordset!codfaccl
        Cad = Cad & " AND fecfaccl = '" & Format(Data1.Recordset!fecfaccl, FormatoFecha) & "' AND numorden =" & Data1.Recordset!numorden
        
        
        
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        Data1.Refresh
        
        
        Ejecuta Cad
        
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
Dim Cad As String
Dim impo As Currency
    
    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    
    If SePuedeEliminar2 < 3 Then Exit Sub
    
    

    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    If impo < 0 Then
        MsgBox "Los abonos no se realizan por caja", vbExclamation
        Exit Sub
    End If


    'Mas gastos
    If Text1(38).Text <> "" Then impo = impo + ImporteFormateado(Text1(38).Text)
    'Menos ya pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)
    
    If impo <= 0 Then
        MsgBox "Totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
    'Devolvera muuuuchas cosas
    'serie factura fecfac numvto
    Cad = Text1(13).Text & "|" & Format(Text1(1).Text, "0000000") & "|" & Text1(2).Text & "|" & Text1(3).Text & "|"
    'Codmacta nommacta codforpa   nomforpa   importe
    Cad = Cad & Text1(4).Text & "|" & Text2(0).Text & "|" & Text1(0).Text & "|" & Text2(1).Text & "|" & CStr(impo) & "|"
    'Lo que lleva cobrado
    Cad = Cad & Text1(8).Text & "|"
    
    
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
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
        .Buttons(10).Image = 10
        '.Buttons(11).Image = 16
        .Buttons(12).Image = 27
        .Buttons(13).Image = 29
        .Buttons(15).Image = 15
        
    
        .Buttons(17).Image = 6
        .Buttons(18).Image = 7
        .Buttons(19).Image = 8
        .Buttons(20).Image = 9
        
    End With
    
    Me.SSTab1.TabVisible(1) = False
    'Cago los iconos
    CargaImagenesAyudas Me.imgCuentas, 1, "Buscar cuenta"
    CargaImagenesAyudas Me.imgFecha, 2
    Carga1ImagenAyuda Me.imgDepart, 1
    Carga1ImagenAyuda imgSerie, 1
    Carga1ImagenAyuda Me.imgAgente, 1
    Me.SSTab1.Tab = 0
    Me.Icon = frmPpal.Icon
    LimpiarCampos
    FrameSeguro.Visible = vParam.TieneOperacionesAseguradas
    
    'Recaudacion ejecutiva
    Label1(27).Visible = vParam.RecaudacionEjecutiva
    Text1(48).Visible = vParam.RecaudacionEjecutiva
    imgFecha(7).Visible = vParam.RecaudacionEjecutiva
    
    
    
    '## A mano
    NombreTabla = "scobro"
    Ordenacion = " ORDER BY numserie,codfaccl,fecfaccl,numorden"
        
    PonerOpcionesMenu
    
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
    End If

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    txtPendiente.Text = ""
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
    cboTipoRem.ListIndex = -1
    lblAsegurado.Visible = False
    lblIndicador.Caption = ""
End Sub



Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Text1(34).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim Cad As String

    If CadenaDevuelta <> "" Then
        If DevfrmCCtas <> "" Then
    
            HaDevueltoDatos = True
            DevfrmCCtas = CadenaDevuelta
            
        Else
                HaDevueltoDatos = True
                Screen.MousePointer = vbHourglass
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
                Cad = DevfrmCCtas
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                Cad = Cad & " AND " & DevfrmCCtas
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
                Cad = Cad & " AND " & DevfrmCCtas
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
                Cad = Cad & " AND " & DevfrmCCtas
                DevfrmCCtas = Cad
                If DevfrmCCtas = "" Then Exit Sub
                '   Como la clave principal es unica, con poner el sql apuntando
                '   al valor devuelto sobre la clave ppal es suficiente
                'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                'If CadB <> "" Then CadB = CadB & " AND "
                'CadB = CadB & Aux
                'Se muestran en el mismo form
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & DevfrmCCtas & " " & Ordenacion
                PonerCadenaBusqueda
                Screen.MousePointer = vbDefault
        End If
    Else
        DevfrmCCtas = ""
    End If
End Sub

Private Sub PonerDatoDevuelto(CadenaDevuelta As String)
Dim Cad As String
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(13), CadenaDevuelta, 1)
                Cad = DevfrmCCtas
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                Cad = Cad & " AND " & DevfrmCCtas
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
                Cad = Cad & " AND " & DevfrmCCtas
                DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
                Cad = Cad & " AND " & DevfrmCCtas
                DevfrmCCtas = Cad
                If DevfrmCCtas = "" Then Exit Sub
                '   Como la clave principal es unica, con poner el sql apuntando
                '   al valor devuelto sobre la clave ppal es suficiente
                'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                'If CadB <> "" Then CadB = CadB & " AND "
                'CadB = CadB & Aux
                'Se muestran en el mismo form
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & DevfrmCCtas & " " & Ordenacion
                PonerCadenaBusqueda
                Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    Text1(33).Text = RecuperaValor(CadenaSeleccion, 3)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
       Text1(0) = RecuperaValor(CadenaSeleccion, 1)
       Text2(1) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgAgente_Click()
    Set frmA = New frmAgentes
    frmA.DatosADevolverBusqueda = "0|1|"
    frmA.Show vbModal
    Set frmA = Nothing
    
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim Cad As String
Dim Z
    Screen.MousePointer = vbHourglass
    If Index = 1 Then
'    DevfrmCCtas = "0"
'    Cad = "Código|codforpa|N|20·"
'    Cad = Cad & "Descripción|nomforpa|T|60·"
'    Cad = Cad & "SIGLAS|Siglas|T|20·"
'    Set frmB = New frmBuscaGrid
'    frmB.vCampos = Cad
'    frmB.vTabla = "sforpa"
'    frmB.vSQL = ""
'    HaDevueltoDatos = False
'    '###A mano
'    frmB.vDevuelve = "0|1|"
'    frmB.vTitulo = "Formas de pago"
'    frmB.vSelElem = 0
'    '#
'    frmB.Show vbModal
'    Set frmB = Nothing
'    If DevfrmCCtas <> "" Then
'       Text1(0) = RecuperaValor(DevfrmCCtas, 1)
'       Text2(1) = RecuperaValor(DevfrmCCtas, 2)
'    End If
        
        Set frmF = New frmFormaPago
        frmF.DatosADevolverBusqueda = "0|"
        frmF.Show vbModal
        Set frmF = Nothing
    
        
    
    Else
        'Cuentas
        imgFecha(0).Tag = Index
        Set frmCCtas = New frmColCtas
        DevfrmCCtas = ""
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        If DevfrmCCtas <> "" Then
            If Index = 0 Then
                Text1(4 + Index) = RecuperaValor(DevfrmCCtas, 1)
            Else
                Text1(7 + Index) = RecuperaValor(DevfrmCCtas, 1)
            End If

            Text2(Index).Text = RecuperaValor(DevfrmCCtas, 2)
        End If
    End If
    
End Sub


Private Sub imgDepart_Click()
    If Text1(4).Text = "" Then
        MsgBox "Seleccione la cuenta del cliente.", vbExclamation
        Exit Sub
    End If
    
    Set frmD = New frmDepartamentos
    frmD.vCuenta = Text1(4).Text
    frmD.DatosADevolverBusqueda = "2|3|"
    frmD.Show vbModal
    Set frmD = Nothing
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'En tag pongo el txtfecha asociado
    Select Case Index
    Case 0
        imgFecha(0).Tag = 2
    Case 1
        imgFecha(0).Tag = 5
    Case 2
        imgFecha(0).Tag = 7
    Case 3
        imgFecha(0).Tag = 32
    Case 4, 5, 6
        imgFecha(0).Tag = 41 + Index
    Case 7
        imgFecha(0).Tag = 48
    End Select
    DevfrmCCtas = Format(Now, "dd/mm/yyyy")
    If IsDate(Text1(CInt(imgFecha(0).Tag)).Text) Then _
        DevfrmCCtas = Format(Text1(CInt(imgFecha(0).Tag)).Text, "dd/mm/yyyy")
    Set frmC = New frmCal
    frmC.Fecha = CDate(DevfrmCCtas)
    DevfrmCCtas = ""
    frmC.Show vbModal
    Set frmC = Nothing
    
    
End Sub

Private Sub imgSerie_Click()
    Set frmS = New frmSerie
    frmS.DatosADevolverBusqueda = "S"
    frmS.Show vbModal
    Set frmS = Nothing
End Sub

Private Sub mnBuscar_Click()

    Dim NF As Integer
    Dim Cad As String
    Dim Entidad As String
    Dim BIC As String
    
    Cad = "C:\Documents and Settings\David\Escritorio\bic.txt"
    NF = FreeFile
    Open Cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        
        'sbic(entidad,Nombre,bic)
        Cad = Trim(Cad)
        
        Entidad = Right(Cad, 4)
        Cad = Mid(Cad, 1, Len(Cad) - 4)
        
        BIC = Mid(Cad, 1, 11)
        Cad = Trim(Mid(Cad, 12))
        
        NombreSQL Cad
        Cad = "INSERT INTO sbic(entidad,Nombre,bic) VALUES (" & Entidad & ",'" & Cad & "','" & BIC & "')"
        Conn.Execute Cad
        
        
    Wend
    Close (NF)


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


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo = 1 Then
        'BUSQUEDA
        If KeyCode = 112 Then HacerF1
    ElseIf Modo = 0 Then
        If KeyCode = 27 Then Unload Me
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 44 Then
        'Despues de la fecha prorroga va el btn
        PonerFocoGral Me.cmdAceptar
    Else
        KeyPressGral KeyAscii
    End If
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
    Dim I As Integer
    Dim SQL As String
    Dim Valor
    
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
        
    'Si esta vacio el campo
    If Text1(Index).Text = "" Then
        I = DevuelveText2Relacionado(Index)
        If I >= 0 Then Text2(I).Text = ""
        Exit Sub
    End If
    
    If Not (Index = 4 Or Index = 10 Or Index = 9) Then
        If Modo < 2 Then Exit Sub
    End If
    'Campo con valor
    Select Case Index
    Case 4, 9, 10
            'Cuentas          'Cuentas
            'Cuentas          'Cuentas
        I = DevuelveText2Relacionado(Index)
        DevfrmCCtas = Text1(Index).Text
        If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
            Text1(Index).Text = DevfrmCCtas
            If Modo >= 2 Then Text2(I).Text = SQL
        Else
            If Modo >= 2 Then
                MsgBox SQL, vbExclamation
                Text1(Index).Text = ""
                Ponerfoco Text1(Index)
            End If
            
            Text2(I).Text = ""
            
        End If
        
        'Poner la cuenta bancaria a partir de la cuenta
        If Index = 4 Then Me.lblAsegurado.Visible = False
        If DevfrmCCtas <> "" Then
            If Modo > 2 And Index = 4 Then
                SQL = ""
                For I = 1 To 4
                    SQL = SQL & Text1(27 + I).Text
                Next I
                
        
        
                Valor = DevuelveLaCtaBanco(DevfrmCCtas)
                If Len(Valor) = 5 Then Valor = ""
                If CStr(Valor) <> "" Then
                    If SQL <> "" Then
                        If MsgBox("Poner Cuenta bancaria de la registro del cliente: " & Replace(CStr(Valor), "|", " - ") & "?", vbQuestion + vbYesNo) = vbYes Then SQL = ""
                    End If
                    If SQL = "" Then
                        SQL = DevuelveLaCtaBanco(DevfrmCCtas)
                        For I = 1 To 4
                            Text1(27 + I).Text = RecuperaValor(SQL, I)
                        Next I
                        Text1(49).Text = RecuperaValor(SQL, I)  'I=5
                    End If
                End If
            End If
            If Index = 4 Then
                'Veremos si es asegurado
                If vParam.TieneOperacionesAseguradas Then
                    SQL = DevuelveDesdeBD("numpoliz", "cuentas", "codmacta", DevfrmCCtas, "T")
                    Me.lblAsegurado.Visible = SQL <> ""
                End If
                
                
                If Modo = 3 Then
                    SQL = "concat(if( isnull(forpa),'',forpa),'|',if(isnull(ctabanco),'',ctabanco),'|')"
                    SQL = DevuelveDesdeBD(SQL, "cuentas", "codmacta", DevfrmCCtas, "T")
                    If SQL <> "" Then
                        Text1(0).Text = RecuperaValor(SQL, 1)
                        Text1(9).Text = RecuperaValor(SQL, 2)
                        If Text1(9).Text <> "" Then Text2(2).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(9).Text, "T", Text1(9).Text)
                        If Text1(0).Text <> "" Then Text1_LostFocus 0   'VOLVEMOS A LLAMR a la lostfocus, cuidado con las variables
                    End If
                End If
            End If
            
        End If
     Case 0
        'FORMA DE PAGO
        Text2(1).Tag = ""
        DevfrmCCtas = "tipforpa"
        If Not IsNumeric(Text1(Index).Text) Then
            SQL = "Campo Forma pago debe ser numérico: " & Text1(Index).Text
            MsgBox SQL, vbExclamation
            SQL = ""
        Else
            SQL = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", Text1(Index).Text, "N", DevfrmCCtas)
            If SQL = "" Then
                SQL = "Forma de pago inexistente: " & Text1(Index).Text
                MsgBox SQL, vbExclamation
                SQL = ""
            Else
                Text2(1).Tag = DevfrmCCtas
            End If
        End If
        Text2(1).Text = SQL
        If Text2(1).Tag = "" Then
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
        End If
        
        
    Case 2, 5, 7, 32, 45, 46, 47
        'FECHAS,32
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
        End If
        
    Case 6, 8, 38
        'IMPORTES
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
    Case 3
        'Vencimiento
        'Debe ser numerico
        If Not IsNumeric(Text1(3).Text) Then
            MsgBox "Campo debe ser numerico", vbExclamation
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
        End If
        
    Case 13
        If IsNumeric(Text1(13).Text) Then
            MsgBox "Serie es una letra.", vbExclamation
            Text1(13).Text = ""
            Ponerfoco Text1(13)
        Else
            Text1(13).Text = UCase(Text1(13).Text)
        End If
        
    Case 28 To 31
        'Cuenta bancaria
        If Index < 30 Then
            I = 4
        Else
            If Index = 30 Then
                I = 2
            Else
                I = 10
            End If
        End If
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Cuenta banco debe ser numérico: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
        Else
            'Formateamos
            SQL = ""
            While Len(SQL) <> I
                SQL = SQL & "0"
            Wend
            SQL = SQL & Text1(Index).Text
            Text1(Index).Text = Right(SQL, I)
            
        End If
        
        SQL = ""
        For I = 28 To 31
            SQL = SQL & Text1(I).Text
        Next
        
        If Len(SQL) = 20 And Index = 31 Then 'solo cuando pierde el foco la cuentaban
            'OK. Calculamos el IBAN
            
            
            If Text1(49).Text = "" Then
                'NO ha puesto IBAN
                If DevuelveIBAN2("ES", SQL, SQL) Then Text1(49).Text = "ES" & SQL
            Else
                Valor = CStr(Mid(Text1(49).Text, 1, 2))
                If DevuelveIBAN2(CStr(Valor), SQL, SQL) Then
                    If Mid(Text1(49).Text, 3) <> SQL Then
                        
                        MsgBox "Codigo IBAN distinto del calculado [" & Valor & SQL & "]", vbExclamation
                        'Text1(49).Text = "ES" & SQL
                    End If
                End If
            End If
        End If
        
        
    Case 33
        
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Departamento debe ser numérico: " & Text1(Index).Text, vbExclamation
            I = 0
        Else
            I = 1
            PonerDepartamenteo
            If Text2(4).Text = "" Then I = 0
        End If
        If I = 0 Then
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
            Text2(4).Text = ""
        End If
        
    Case 34
        I = 0
        If Text1(34).Text <> "" Then
            SQL = DevuelveDesdeBD("nombre", "agentes", "codigo", Text1(Index).Text, "N")
            If SQL = "" Then
                MsgBox "No existe el agente: " & Text1(34).Text, vbExclamation
                I = 2
            Else
                I = 1
            End If
        Else
            SQL = ""
        End If
        Text2(5).Text = SQL
        If I = 2 Then Ponerfoco Text1(34)
            
    Case 49
        Text1(Index).Text = UCase(Text1(Index).Text)
    End Select
            
End Sub

Public Function DevuelveText2Relacionado(Index As Integer) As Integer
        DevuelveText2Relacionado = -1
        Select Case Index
        Case 0
            DevuelveText2Relacionado = 1
        Case 4
            DevuelveText2Relacionado = 0
        Case 9
            DevuelveText2Relacionado = 2
        Case 10
            DevuelveText2Relacionado = 3
        End Select
End Function


Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me, BuscaChekc)

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
        
        CadenaDesdeOtroForm = ""
        frmVerCobrosPagos.vSQL = CadB
        frmVerCobrosPagos.OrdenarEfecto = False
        frmVerCobrosPagos.Regresar = True
        frmVerCobrosPagos.Cobros = True
        frmVerCobrosPagos.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            PonerDatoDevuelto CadenaDesdeOtroForm
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
               ' Text1(kCampo).SetFocus
                Ponerfoco Text1(kCampo)
        End If
        
        'Llamamos a al form
'        '##A mano
'        Cad = ""
'        Cad = Cad & ParaGrid(Text1(4), 30, "Proveedor")
'        Cad = Cad & ParaGrid(Text1(1), 30, "Factura")
'        Cad = Cad & ParaGrid(Text1(2), 25, "Fecha")
'        Cad = Cad & ParaGrid(Text1(3), 10, "Numero")
'        If Cad <> "" Then
'            Screen.MousePointer = vbHourglass
'            DevfrmCCtas = ""
'            Set frmB = New frmBuscaGrid
'            frmB.vCampos = Cad
'            frmB.vTabla = NombreTabla
'            frmB.vSQL = CadB
'            HaDevueltoDatos = False
'            '###A mano
'            frmB.vDevuelve = "0|1|2|3|"
'            frmB.vTitulo = "Pagos proveedor"
'            frmB.vSelElem = 0
'            '#
'            frmB.Show vbModal
'            Set frmB = Nothing
'            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                Text1(kCampo).SetFocus
'            End If
'        End If
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
    PonerDepartamenteo
    Text1_LostFocus 34
    Text1_LostFocus 0
    Text3(0).Text = vEmpresa.nomempre
    Text3(1).Text = Text2(0).Text
    
    'SI tiene impagados
    'Para ello la forma de pago debe ser remesa
    'Y tiene que tener el chekc de imagado(contdocu) a 1
    I = 0
    If Text2(1).Tag <> "" Then
        If Val(Text2(1).Tag) = vbTipoPagoRemesa Or Val(Text2(1).Tag) = vbTalon Or Val(Text2(1).Tag) = vbPagare Then
            If Me.Check1(1).Value = 1 Then I = 1
        End If
    End If
    
    PonPendiente
    
    Me.Toolbar1.Buttons(10).Enabled = (I = 1)
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
End Sub


Private Sub PonPendiente()
Dim Importe As Currency

    On Error GoTo EPonPendiente
    'Pendiente
    Importe = Data1.Recordset!impvenci + DBLet(Data1.Recordset!Gastos, "N") - DBLet(Data1.Recordset!impcobro, "N")
    txtPendiente.Text = Format(Importe, FormatoImporte)
    
EPonPendiente:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Err.Clear
    End If
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Dim B As Boolean
    BuscaChekc = ""
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            'Text1(i).BackColor = vbWhite
            Text1(0).BackColor = &H80000018
        Next I
        Text1(28).MaxLength = 4
        Text1(29).MaxLength = 4
        'chkVistaPrevia.Visible = False
    ElseIf Modo = 4 Then
        FrameSeguro.Enabled = True
    End If
    
    'Modo buscar
    If Kmodo = 1 Then
        Text1(28).MaxLength = 0
        Text1(29).MaxLength = 0
    End If
    
    
    Modo = Kmodo
    FrameRemesa.Enabled = Kmodo = 1
    Text1(40).Enabled = Kmodo = 1
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2) And vUsu.Nivel < 2
    
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B
    
    Toolbar1.Buttons(12).Enabled = B
    Toolbar1.Buttons(13).Enabled = B
    
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    DespalzamientoVisible B
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
        'cmdCancelar.Cancel = False
        
    End If
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 2
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
    'Empieza siempre a false
    Toolbar1.Buttons(10).Enabled = False
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = (Modo = 2) Or Modo = 0
    For I = 0 To Text1.Count - 1
        
        Text1(I).Locked = B
        
        If B Then
            Text1(I).BackColor = &H80000018
        Else
            Text1(I).BackColor = vbWhite
        End If
    Next I
    frameContene.Enabled = Not B
    For I = 0 To 6
        If I < 4 Then imgCuentas(I).Visible = Not B
        Me.imgFecha(I).Visible = Not B
    Next I
    Me.imgSerie.Visible = Not B
    Me.imgDepart.Visible = Not B
    Me.imgAgente.Visible = Not B
        
    Text2(1).Tag = ""
    FrameEstaEnCaja.Enabled = (Modo = 1)
    
    
    If Me.FrameRemesa.Enabled Then
        Me.cboTipoRem.BackColor = vbWhite
    Else
        Me.cboTipoRem.BackColor = &H80000018
    End If
        
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim Tipo As Integer

    DatosOk = False
    
    
    DevfrmCCtas = ""
    
    If Text1(34).Text = "" Then
        DevfrmCCtas = vbCrLf & "-  Agente "
        Tipo = 34
    End If
    
    If Text1(9).Text = "" Then
        DevfrmCCtas = DevfrmCCtas & vbCrLf & "-  Cuenta prevista cobro "
        Tipo = 9
    End If
    
    If Text1(4).Text = "" Then
        DevfrmCCtas = DevfrmCCtas & vbCrLf & "-  Cuenta cliente "
        Tipo = 4
    End If
    If DevfrmCCtas <> "" Then
        DevfrmCCtas = "Los siguientes campos son requeridos:" & vbCrLf & vbCrLf & DevfrmCCtas
        MsgBox DevfrmCCtas, vbExclamation
        Ponerfoco Text1(Tipo)
        Exit Function
    End If
    
    Text2(1).Tag = ""
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'NUmero serie
    DevfrmCCtas = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", Text1(13).Text, "T")
    If DevfrmCCtas = "" Then
        B = False
        MsgBox "Serie no existe", vbExclamation
        Exit Function
    End If
    
    
    
    DevfrmCCtas = DevuelveDesdeBD("tipforpa", "sforpa", "codforpa", Text1(0).Text, "N")
    Tipo = CInt(DevfrmCCtas)
    

    
    DevfrmCCtas = Trim(Text1(28).Text) & Trim(Text1(29).Text) & Trim(Text1(31).Text)
    
    
    
    
    'Para preguntar por el Banco
    B = False
    If DevfrmCCtas <> "" Then
        If Val(DevfrmCCtas) <> 0 Then B = True
    End If
        
    If B Then
        'Vale, hay campos y son numericos
        'La cuenta contable si digi control, si tiene valor, tiene que ser longitud 18
        If Len(DevfrmCCtas) < 18 Then
            MsgBox "Cuenta bancaria incorrecta", vbExclamation
            Exit Function
        End If
    End If
        
        
    If B Then
            BuscaChekc = CodigoDeControl(DevfrmCCtas)
            If BuscaChekc <> Text1(30).Text Then
                BuscaChekc = vbCrLf & "Código de control calculado: " & BuscaChekc & vbCrLf
                BuscaChekc = "Error en la cuenta contable: " & vbCrLf & BuscaChekc & vbCrLf & "Codigo de control: " & Text1(30).Text & vbCrLf & vbCrLf
                
                BuscaChekc = BuscaChekc & "Desea continuar?"
                If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
            'Compruebo EL IBAN
            'Meto el CC
            DevfrmCCtas = Mid(DevfrmCCtas, 1, 8) & Me.Text1(30).Text & Mid(DevfrmCCtas, 9)
            BuscaChekc = ""
            If Me.Text1(49).Text <> "" Then BuscaChekc = Mid(Text1(49).Text, 1, 2)
                
            If DevuelveIBAN2(BuscaChekc, DevfrmCCtas, DevfrmCCtas) Then
                If Me.Text1(49).Text = "" Then
                    If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(49).Text = BuscaChekc & DevfrmCCtas
                Else
                    If Mid(Text1(49).Text, 3) <> DevfrmCCtas Then
                        DevfrmCCtas = "Calculado : " & BuscaChekc & DevfrmCCtas
                        DevfrmCCtas = "Introducido: " & Me.Text1(49).Text & vbCrLf & DevfrmCCtas & vbCrLf
                        DevfrmCCtas = "Error en codigo IBAN" & vbCrLf & DevfrmCCtas & "Continuar?"
                        If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
                    End If
                End If
            End If
            
            
    Else
        If Tipo = vbTipoPagoRemesa Then
                DevfrmCCtas = "Debe poner cuenta bancaria. Desea continuar?"
                If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
    End If
    
   
        If Modo = 4 Then
            If DBLet(Me.Data1.Recordset!recedocu, "N") = 1 Then
                'Tiene la marca de documento recibido
                'Veremos si se la ha quitado
                If Me.Check1(0).Value = 0 Then
                    DevfrmCCtas = "Seguro que desea quitarle la marca de documento recibido?"
                    If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
        End If

    
    'Nuevo. 12 Mayo 2008
    B = CuentaBloqeada(Me.Text1(4).Text, CDate(Text1(2).Text), True)
    If B Then
        If (vUsu.Codigo Mod 100) > 0 Then Exit Function
    End If
    
    
    
    'Ultimas comprobaciones
    If vParam.TieneOperacionesAseguradas Then
        B = Me.Text1(45).Text <> "" Or Me.Text1(46).Text <> "" Or Me.Text1(47).Text <> ""
        If B Then
            'Tiene valores en fechas de riesgo/aviso/siniestro
            If Me.lblAsegurado.Visible Then
                'ok. el cliente tiene operaciones aseguradas
                
            Else
                MsgBox "No debe indicar fechas de operaciones aseguradas" & vbCrLf & "-Falta pago/prorroga/aviso siniestro" & vbCrLf & " Si no esta asegurado", vbExclamation
                Ponerfoco Me.Text1(45)
                Exit Function
            End If
        End If
    End If
    
    
    DatosOk = True
End Function




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    BotonBuscar
Case 2
    BotonVerTodos
Case 6
    BotonAnyadir
Case 7
     BotonModificar
'    If BLOQUEADesdeFormulario(Me) Then BotonModificar
Case 8
    BotonEliminar
    
Case 10
    'Ver los impagados
    If Text1(13).Text = "" Then Exit Sub
    
    CadenaDesdeOtroForm = Text1(13).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|"
    frmVarios.Opcion = 10
    frmVarios.Show vbModal
Case 12
    'Cobros parciales
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Modo <> 2 Then Exit Sub
    If Text2(1).Tag <> "" Then
        'If Val(Text2(1).Tag) < 4 Or Val(Text2(1).Tag) > 5 Then 'El 4 y el 5 son recibo bancario y confirming
        If Val(Text2(1).Tag) <> vbTipoPagoRemesa Then
            
            If SePuedeEliminar2 < 3 Then Exit Sub
        
            'Bloqueamos
            If BloqueaRegistroForm(Me) Then
                RealizarPagoCuenta
                DesBloqueaRegistroForm Text1(0)
            End If
        Else
            'MsgBox "Lo pagos a cuenta no se realizan sobre RECIBOS y CONFIRMING", vbExclamation
            MsgBox "Lo pagos a cuenta no se realizan sobre RECIBOS BANCARIOS", vbExclamation
        End If
    End If
    
    
Case 13
    DividirVencimiento
    
Case 15
    mnSalir_Click


Case 17 To 20
    Desplazamiento Button.Index - 17
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 17 To 20
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub


Private Sub PonerCtasIVA()
On Error GoTo EPonerCtasIVA

    Text1_LostFocus 4
    Text1_LostFocus 0
    Text1_LostFocus 9
    Text1_LostFocus 10
Exit Sub
EPonerCtasIVA:
    MuestraError Err.Number, "Poniendo valores ctas. IVA", Err.Description
End Sub



Private Sub Ponerfoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub


'Si no esta en transferencia o en una remesa
'entonces dejare que modifique algun dato basico
'Realmente solo la cta bancaria
Private Function SePuedeEliminar2() As Byte


    SePuedeEliminar2 = 0 'NO se puede eliminar

    SePuedeEliminar2 = 1
    If Val(DBLet(Data1.Recordset!CodRem)) > 0 Then
        MsgBox "Pertenece a una remesa", vbExclamation
        'Noviembre 2009
        If vUsu.Nivel < 2 Then
            If CStr(Data1.Recordset!siturem) = "Q" Or CStr(Data1.Recordset!siturem) = "Y" Then
                'DEJO ELIMINARLO
                If MsgBox("Efecto remesado. Situacion: " & Data1.Recordset!siturem & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
                espera 1
                If MsgBox("¿Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Else
                'Tampoco dejamos continuar
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    'Si no esta en transferencia
    If Val(DBLet(Data1.Recordset!transfer)) > 0 Then
        MsgBox "Pertenece a una transferencia", vbExclamation
        Exit Function
    End If
    
    
    'SI no esta en la caja
    If Val(DBLet(Data1.Recordset!estacaja)) > 0 Then
        MsgBox "Esta en caja. ", vbExclamation
        Exit Function
    End If
    
    'Si  tiene documento recibido
    If Val(DBLet(Data1.Recordset!recedocu)) > 0 Then
        'Documento recibido
        '
        DevfrmCCtas = "numserie='" & Data1.Recordset!NUmSerie
        DevfrmCCtas = DevfrmCCtas & "' AND fecfaccl='" & Format(Data1.Recordset!fecfaccl, FormatoFecha)
        DevfrmCCtas = DevfrmCCtas & "' AND numfaccl=" & Data1.Recordset!codfaccl
        DevfrmCCtas = DevfrmCCtas & " AND numvenci"
        DevfrmCCtas = DevuelveDesdeBD("id", "slirecepdoc", DevfrmCCtas, Data1.Recordset!numorden)
        If DevfrmCCtas <> "" Then
            DevfrmCCtas = "Esta en la recepcion de documentos. Numero: " & DevfrmCCtas
            MsgBox DevfrmCCtas, vbExclamation
            DevfrmCCtas = ""
            Exit Function
        End If
    End If
    
    
    
    
    
    SePuedeEliminar2 = 3  'SI SE PUEDE ELIMINAR

    Screen.MousePointer = vbDefault
End Function


Private Sub PonerDepartamenteo()
Dim C As String
Dim O As Boolean

    O = False
    
    If Text1(4).Text <> "" Then
        If Text1(33).Text <> "" Then
                    
            Set miRsAux = New ADODB.Recordset
            C = "Select Descripcion FROM Departamentos WHERE codmacta ='" & Text1(4).Text
            C = C & "' AND Dpto =" & Text1(33).Text
            miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux.Fields(0)) Then
                    C = miRsAux.Fields(0)
                    O = True
                End If
            End If
            miRsAux.Close
            Set miRsAux = Nothing
        End If
    End If
    If O Then
        Text2(4).Text = C
    Else
        Text2(4).Text = ""
    End If
    
End Sub
    



Private Sub RealizarPagoCuenta()
Dim impo As Currency
    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    'Gastos
    If Text1(38).Text <> "" Then impo = impo + ImporteFormateado(Text1(38).Text)
    'Pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)
    
    'Si impo>0 entonces TODAVIA puedn pagarme algo
    If impo = 0 Then
        'Cosa rara. Esta todo el importe pagado
        Exit Sub
    End If
        
    frmParciales.Cobro = True
    frmParciales.Vto = Text1(13).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text1(5).Text & "|"
    frmParciales.Importes = Text1(6).Text & "|" & Text1(38).Text & "|" & Text1(8).Text & "|"
    frmParciales.Cta = Text1(4).Text & "|" & Text2(0).Text & "|" & Text1(9).Text & "|" & Text2(2).Text & "|"
    frmParciales.FormaPago = Val(Text2(1).Tag)
    frmParciales.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'Hay que refrescar los datos
        lblIndicador.Caption = ""
        If SituarData Then
            
            PonerCampos
            
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
End Sub

Private Sub HacerF1()
Dim C As String
    
    C = ObtenerBusqueda(Me, BuscaChekc)
    If C = "" Then Text1(13).Text = "*"  'Para que busqu toooodo
    cmdAceptar_Click
End Sub




Private Sub DividirVencimiento()
Dim Im As Currency

    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si esta totalmente cobrado pues no podemos desdoblar ekl vto
    
    
    
    If Val(DBLet(Data1.Recordset!transfer, "N")) = 1 Then
        MsgBox "Pertenece a una transferencia", vbExclamation
        Exit Sub
    End If
    If Val(Data1.Recordset!estacaja) = 1 Then
        MsgBox "Esta en caja", vbExclamation
        Exit Sub
    End If
    
    
    Im = Data1.Recordset!impvenci + DBLet(Data1.Recordset!Gastos, "N")
    Im = Im - DBLet(Data1.Recordset!impcobro, "N")
    If Im = 0 Then
        MsgBox "NO puede dividir el vencimiento. Importe totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
    
       'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
    
    CadenaDesdeOtroForm = "numserie = '" & Data1.Recordset!NUmSerie & "' AND codfaccl = " & Data1.Recordset!codfaccl
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND fecfaccl = '" & Format(Data1.Recordset!fecfaccl, FormatoFecha) & "'|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Data1.Recordset!numorden & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CStr(Im) & "|"
    
    
    'Ok, Ahora pongo los labels
    frmListado.Opcion = 27
    frmListado.Label4(56).Caption = Text2(0).Text
    frmListado.Label4(57).Caption = Data1.Recordset!NUmSerie & Format(Data1.Recordset!codfaccl, "000000") & " / " & Data1.Recordset!numorden & "      de " & Format(Data1.Recordset!fecfaccl, "dd/mm/yyyy")
    
    'Si ya ha cobrado algo...
    Im = DBLet(Data1.Recordset!impcobro, "N")
    If Im > 0 Then frmListado.txtImporte(1).Text = txtPendiente.Text
    
    frmListado.Show vbModal
    If CadenaDesdeOtroForm <> "" Then

            CadenaConsulta = "numserie = '" & Data1.Recordset!NUmSerie & "' AND codfaccl = " & Data1.Recordset!codfaccl
            CadenaConsulta = CadenaConsulta & " AND fecfaccl = '" & Format(Data1.Recordset!fecfaccl, FormatoFecha) & "'"
            CadenaConsulta = "Select * from scobro WHERE " & CadenaConsulta
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            If Data1.Recordset.RecordCount <= 0 Then
                   MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
            Else
                DevfrmCCtas = ""
                While DevfrmCCtas = ""
                    If CStr(Data1.Recordset!numorden) = CadenaDesdeOtroForm Then
                        DevfrmCCtas = "YA"
                    Else
                        If Data1.Recordset.EOF Then
                            DevfrmCCtas = "EOF"
                        Else
                            Data1.Recordset.MoveNext
                        End If
                    End If
                Wend
                If DevfrmCCtas = "EOF" Then Data1.Recordset.MoveFirst
                PonerCampos
            End If
    End If
End Sub
