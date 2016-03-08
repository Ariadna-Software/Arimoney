VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRemesaAutomatica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesa Automática"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmRemesaAutomatica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8880
   Begin VB.CommandButton CmdSalir 
      Height          =   460
      Left            =   8250
      Picture         =   "frmRemesaAutomatica.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Salir"
      Top             =   5700
      Width           =   495
   End
   Begin VB.CommandButton cmdVer 
      Height          =   460
      Left            =   6360
      Picture         =   "frmRemesaAutomatica.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Pedir datos"
      Top             =   5700
      Width           =   495
   End
   Begin VB.CommandButton CmdAcep 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6495
      TabIndex        =   53
      Top             =   5745
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7665
      TabIndex        =   52
      Top             =   5745
      Width           =   1035
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   7470
      MaxLength       =   6
      TabIndex        =   51
      Top             =   5400
      Width           =   870
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   465
      Left            =   7350
      Picture         =   "frmRemesaAutomatica.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Eliminar Efecto"
      Top             =   5700
      Width           =   495
   End
   Begin VB.CommandButton CmdCambiarBanco 
      Height          =   465
      Left            =   6855
      Picture         =   "frmRemesaAutomatica.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Cambiar Banco"
      Top             =   5700
      Width           =   480
   End
   Begin VB.Frame FrameCalculados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2070
      Left            =   120
      TabIndex        =   24
      Top             =   1350
      Width           =   8610
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   4680
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "1234"
         Top             =   540
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   225
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   4680
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   900
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   225
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   4680
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   225
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   9
         Left            =   4680
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1620
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   900
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   14
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   19
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   1620
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   8
         Left            =   225
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "1234"
         Top             =   900
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   10
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "1234"
         Top             =   1260
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   15
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "1234"
         Top             =   1620
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "1234"
         Top             =   540
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "1234"
         Top             =   540
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "1234567890"
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "1234"
         Top             =   900
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "1234"
         Top             =   1260
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   16
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "1234"
         Top             =   1620
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "1234"
         Top             =   900
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "1234"
         Top             =   1260
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   17
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "1234"
         Top             =   1620
         Width           =   585
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "1234567890"
         Top             =   900
         Width           =   1305
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   13
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "1234567890"
         Top             =   1260
         Width           =   1305
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   18
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1234567890"
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "Importe Remesado"
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
         Left            =   6525
         TabIndex        =   47
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label Label7 
         Caption         =   "Importe a Remesar"
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
         Left            =   4680
         TabIndex        =   46
         Top             =   270
         Width           =   1620
      End
      Begin VB.Label Label10 
         Caption         =   "Banco de la Remesa"
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
         Height          =   240
         Left            =   225
         TabIndex        =   45
         Top             =   225
         Width           =   3765
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   810
         Picture         =   "frmRemesaAutomatica.frx":162A
         Top             =   540
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   810
         Picture         =   "frmRemesaAutomatica.frx":172C
         Top             =   945
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   810
         Picture         =   "frmRemesaAutomatica.frx":182E
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   810
         Picture         =   "frmRemesaAutomatica.frx":1930
         Top             =   1665
         Width           =   240
      End
   End
   Begin VB.Frame FrameIntro 
      Caption         =   "Datos Remesa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1335
      Left            =   135
      TabIndex        =   15
      Top             =   30
      Width           =   8595
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7515
         TabIndex        =   6
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox txtFec 
         Height          =   315
         Index           =   1
         Left            =   4185
         TabIndex        =   2
         Tag             =   "La fecha de factura"
         Text            =   "99/99/9999"
         Top             =   585
         Width           =   1035
      End
      Begin VB.TextBox txtFec 
         Height          =   315
         Index           =   2
         Left            =   4185
         TabIndex        =   3
         Tag             =   "La fecha de recepción"
         Text            =   "99/99/9999"
         Top             =   945
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   6165
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "El número de factura "
         Text            =   "Text1"
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox txtFec 
         Height          =   315
         Index           =   0
         Left            =   1890
         TabIndex        =   0
         Tag             =   "La fecha de factura"
         Text            =   "99/99/9999"
         Top             =   450
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   6165
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "El número de factura "
         Text            =   "Text1"
         Top             =   585
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmRemesaAutomatica.frx":1A32
         Left            =   1755
         List            =   "frmRemesaAutomatica.frx":1A34
         TabIndex        =   1
         Tag             =   "Tipo de Pago|N|N|||sefect|tipoforp|||"
         Text            =   "Combo"
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Pago"
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
         Height          =   255
         Index           =   2
         Left            =   225
         TabIndex        =   23
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Vencimientos"
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
         Height          =   255
         Index           =   0
         Left            =   3150
         TabIndex        =   22
         Top             =   315
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   21
         Top             =   990
         Width           =   540
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   3960
         Picture         =   "frmRemesaAutomatica.frx":1A36
         Top             =   630
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   3960
         Picture         =   "frmRemesaAutomatica.frx":1B38
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   5625
         TabIndex        =   20
         Top             =   990
         Width           =   555
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   1575
         Picture         =   "frmRemesaAutomatica.frx":1C3A
         Top             =   495
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Remesa"
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
         Height          =   255
         Index           =   8
         Left            =   225
         TabIndex        =   19
         Top             =   495
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   9
         Left            =   3420
         TabIndex        =   18
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   5625
         TabIndex        =   17
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Efectos"
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
         Height          =   255
         Index           =   11
         Left            =   5355
         TabIndex        =   16
         Top             =   315
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRemesaAutomatica.frx":1D3C
      Height          =   2235
      Left            =   120
      TabIndex        =   48
      Top             =   3420
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   3942
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   5940
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
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
      UserName        =   "root"
      Password        =   "aritel"
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
End
Attribute VB_Name = "frmRemesaAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBPr As frmBancosPropios
Attribute frmBPr.VB_VarHelpID = -1

Dim PrimeraVez As Boolean

Dim SQL As String
Dim RS As Adodb.Recordset

Dim Importe As Currency
Dim i As Integer
Dim PrimeraSeleccion As Boolean
Dim ClickAnterior As Byte '0 Empezar 1.-Debe 2.-Haber
Dim VerAlbaranes As Boolean
'Con estas dos variables
Dim ContadorBus As Integer
Dim Checkear As Boolean

Dim NumRemes As Long

Dim Importe1 As Currency
Dim Importe2 As Currency
Dim Importe3 As Currency
Dim Importe4 As Currency
Dim TotalEfectos As Currency
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim mC As Contadores

Private Sub CmdAcep_Click()
    If DatosBanOk Then
            ModificarRegistro
            PonerModo 2
    Else
        MsgBox "Código de banco incorrecto.", vbExclamation
        txtAux(0).Text = ""
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub cmdAceptar_Click()
    If Datos1Ok Then
       If Not BloqueoManual(True, "REMESA", Combo1.Text) Then
            MsgBox "No se puede remesar ese tipo de pago. Hay alguien remesándolo.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
       End If

        PonerModo 1
        InicializarImportes
    Else
        MsgBox "No existe datos entre esos límites. Reintroduzca.", vbExclamation
        PonerModo 0
    End If
End Sub

Private Sub cmdCancel_Click()
    
    PonerModo 2
    
End Sub

Private Sub cmdEliminar_Click()
Dim SQL As String

On Error GoTo eEliminarRegistro
    
    SQL = SQL & "Va a eliminar el efecto: "
    SQL = SQL & Adodc1.Recordset.Fields(0).Value & " - " & Adodc1.Recordset.Fields(1).Value
    SQL = SQL & " - " & Adodc1.Recordset.Fields(2).Value & " - " & Adodc1.Recordset.Fields(3).Value
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? " 'VRS:1.0.1(11)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
            SQL = " delete from sremes where tipofact = "
            SQL = SQL & Adodc1.Recordset.Fields(0).Value & " and numserie = '"
            SQL = SQL & DevNombreSQL(Adodc1.Recordset.Fields(1).Value) & "' and numfactu = "
            SQL = SQL & Adodc1.Recordset.Fields(2).Value & " and ordefect = "
            SQL = SQL & Adodc1.Recordset.Fields(3).Value
            
            Conn.Execute SQL
            
            RecalculoImportes
     End If
    
eEliminarRegistro:
    If Err.Number <> 0 Then
        MuestraError 0, "Error eliminando registro."
    Else
        PonerModo 2
    End If
End Sub

Private Sub cmdVer_Click()
       
    If DatosOk Then
        InsertaRemesa
        PonerModo 2
    End If
 
'    'Desbloqueamos
'    FrameIntro.Enabled = True
'    'fecha de recepcion now
'    txtFec(1).Text = Format(Now, "dd/mm/yyyy")
'
'    cmdAceptar.Enabled = True
'    cmdAceptar.Visible = True
'
'    PonerFoco Text1(0)
'
'    BloqueoManual False, "scaalp", "Borrar"
'    BloqueoManual False, "RECFAC", Text1(1).Text
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdCambiarBanco_Click()
Dim i As Integer
Dim SQL As String
Dim v_count As Integer
Dim Cad As String
Dim anc As Single
    
    
    
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    txtAux(0).Text = Adodc1.Recordset.Fields!banremes
    
    LLamaLineas anc, 2, False
    PonerFoco txtAux(0)
    PonerModo 3
End Sub
Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub
    
Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim b As Boolean
    DeseleccionaGrid
    
    DataGrid1.Enabled = True
    
    txtAux(0).Top = alto
    txtAux(0).Text = ""
End Sub
    
'Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
'    Dim i As Integer
'    Dim J As Integer
'
'    DataGrid1.Enabled = Not Visible
'
'    J = Txtaux.Count
'    For i = 0 To J - 1
'        If i <> 6 Then
'            Txtaux(i).Visible = Visible
'            Txtaux(i).Top = Altura
'        End If
'    Next i
'
'    cmdAux(0).Visible = Visible
'    cmdAux(0).Top = Altura
'    cmdAux(1).Visible = Visible
'    cmdAux(1).Top = Altura
'
'    If Limpiar Then
'        For i = 0 To J - 1
'            Txtaux(i).Text = ""
'        Next i
'    End If
'
'End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
 '       cmdPedirDatos_Click
    End If
End Sub

Private Sub Form_Load()
    
    Me.Top = 0
    Me.Left = 0
    
'    cmdVer.Picture = frmPpal.ImgListComun.ListImages(18).Picture
'    CmdCambiarBanco.Picture = frmPpal.ImgListComun.ListImages(4).Picture
'    cmdEliminar.Picture = frmPpal.ImgListComun.ListImages.Item(5).Picture
'    cmdsalir.Picture = frmPpal.ImgListComun.ListImages.Item(15).Picture
    
    CargarCombo1
    
    PonerModo 0
    
    NumRemes = -1
    
    PrimeraVez = True
End Sub

Private Sub PonerModo(modo As Byte)

    FrameIntro.Enabled = (modo = 0)
    FrameCalculados.Enabled = (modo = 1)
    DataGrid1.Enabled = (modo = 2)
    cmdVer.Enabled = (modo = 1)
    CmdCambiarBanco.Enabled = (modo = 2)
    cmdEliminar.Enabled = (modo = 2)
    CmdSalir.Enabled = (modo < 3)
    
    cmdVer.Visible = (modo = 1)
    CmdCambiarBanco.Visible = (modo = 2)
    cmdEliminar.Visible = (modo = 2)
    CmdSalir.Visible = (modo < 3)
    
    CmdAcep.Visible = (modo = 3)
    cmdCancel.Visible = (modo = 3)
    CmdAcep.Enabled = (modo = 3)
    cmdCancel.Enabled = (modo = 3)
    txtAux(0).Enabled = (modo = 3)
    txtAux(0).Visible = (modo = 3)
    
    
    
    Select Case modo
        Case 0
            Limpiar Me
            txtFec(0).Text = Format(Now, "dd/mm/yyyy")
            PonerFoco txtFec(0)
             CargaGrid (modo = 2)
        Case 1
            PonerFoco Text1(2)
             CargaGrid (modo = 2)
        Case 2
            DataGrid1.SetFocus
            CargaGrid (modo = 2)
        Case 3
            PonerFoco txtAux(0)
    End Select
    
   
    
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Adodc1.Recordset.EOF Then
        ' desactualizamos contador
        If NumRemes > 0 Then
             mC.DevolverContador "1", 5, NumRemes
        End If
    Else
        If MsgBox("Desea actualizar los datos de la remesa", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
            ' borramos la remesa
            Conn.Execute "delete from sremes where numremes = " & NumRemes
            ' desactualizamos contador
             mC.DevolverContador "1", 5, NumRemes
        Else
             ActualizaCartera
        End If
    End If
    
    Set mC = Nothing
    
    'Desbloqueamos
    BloqueoManual False, "REMESA", Combo1.Text
'    If vParam.HayContabilidad Then ConnConta.Close
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFec(CInt(txtFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

'Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
'    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub


Private Sub imgppal_Click(Index As Integer)
    
    Select Case Index
        Case 0, 1
            Set frmC = New frmCal
            frmC.Fecha = Now
            txtFec(0).Tag = Index
            If txtFec(Index).Text <> "" Then
                If IsDate(txtFec(Index).Text) Then frmC.Fecha = CDate(txtFec(0).Text)
            End If
            frmC.Show vbModal
            Set frmC = Nothing
        Case 2
'            Set frmProv = New frmProveedores
'            frmProv.DatosADevolverBusqueda = "0|1|"
'            frmProv.Show
    End Select
End Sub




Private Sub txtAux_GotFocus(Index As Integer)
    txtAux(Index).SelStart = 0
    txtAux(Index).SelLength = Len(Text1(Index).Text)

End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim SQL As String
    If txtAux(0).Text = "" Then
        MsgBox "Debe introducir un valor en el código de banco. Reintroduzca.", vbExclamation
        PonerFoco txtAux(Index)
    Else
        SQL = ""
        If EsNumerico(txtAux(Index).Text) Then
            SQL = DevuelveDesdeBD(1, "nombanpr", "sbanpr", "codbanpr|", txtAux(0).Text & "|", "N|", 1)
            If SQL = "" Then
                MsgBox "Código de banco no existe.Reintroduzca.", vbExclamation
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            Else
'                Txtaux(Index).Text = Format(Txtaux(Index).Text, "0000")
                PonerFoco CmdAcep
            End If
        End If
    End If

End Sub

Private Sub txtfec_GotFocus(Index As Integer)
    txtFec(Index).SelStart = 0
    txtFec(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub txtfec_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfec_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    Mal = True
    If txtFec(Index).Text = "" Then Exit Sub
        
    If Not EsFechaOK(txtFec(Index)) Then
        MsgBox "No es una fecha correcta", vbExclamation
    Else
        Mal = False
    End If
    
    If Mal Then PonerFoco txtFec(Index)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim valor As Currency

    Text1(Index).Text = Trim(Text1(Index).Text)
    BloqueoCampos (Index)
    If Text1(Index).Text = "" Then Exit Sub
    Select Case Index
        Case 0
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ningún campo de texto", vbExclamation
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, ">")
        Case 1
            If EsNumerico(Text1(Index).Text) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
'                Text2.Text = DevuelveDesdeBD(1, "nomprove", "sprove", "codprove|", Text1(Index).Text & "|", "N|", 1)
'                If Text2.Text = "" Then
'                    MsgBox "Coódigo no existe. Reintroduzca", vbExclamation
'                    PonerFoco Text1(Index)
'                End If
            Else
                PonerFoco Text1(Index)
            End If
        
        Case 2, 4, 6, 8
            If EsNumerico(Text1(Index).Text) Then
                 If DatosBanco(Text1(Index).Text, Index) Then
                    Text1(Index + 1).Enabled = True
                    PonerFoco Text1(Index + 1)
                 Else
                    Text1(Index + 1).Enabled = False
                 End If
            End If
        Case 3, 5, 7, 9
            If EsNumerico(Text1(Index)) Then
                If InStr(1, Text1(Index).Text, ",") > 0 Then
                        valor = ImporteFormateado(Text1(Index).Text)
                Else
                        valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                End If
                Select Case Index
                    Case 3
                        Importe1 = valor
                        Text1(Index).Text = Format(Importe1, "###,###,##0.00")
                    Case 5
                        Importe2 = valor
                        Text1(Index).Text = Format(Importe2, "###,###,##0.00")
                    Case 7
                        Importe3 = valor
                        Text1(Index).Text = Format(Importe3, "###,###,##0.00")
                    Case 9
                        Importe4 = valor
                        Text1(Index).Text = Format(Importe4, "###,###,##0.00")
                End Select
            End If
        
    End Select

End Sub

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub CargaGrid(Enlaza As Boolean)
Dim b As Boolean
    b = DataGrid1.Enabled
    DataGrid1.Enabled = False
    CargaGrid2 Enlaza
    DataGrid1.Enabled = b
End Sub

Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(Enlaza)
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockPessimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    
    'Cuenta
    DataGrid1.Columns(0).Caption = "T"
    DataGrid1.Columns(0).Width = 250
    DataGrid1.Columns(0).NumberFormat = "0"
    
    DataGrid1.Columns(1).Caption = "Serie"
    DataGrid1.Columns(1).Width = 550
    
    DataGrid1.Columns(2).Caption = "Factura"
    DataGrid1.Columns(2).Width = 800
    DataGrid1.Columns(2).NumberFormat = "0000000"
    
    DataGrid1.Columns(3).Caption = "Ord."
    DataGrid1.Columns(3).Width = 500
    DataGrid1.Columns(3).NumberFormat = "0"
    DataGrid1.Columns(3).Alignment = dbgCenter
    
    DataGrid1.Columns(4).Caption = "F.Efecto"
    DataGrid1.Columns(4).Width = 1000 '940 '4395

    DataGrid1.Columns(5).Caption = "Socio/Cliente"
    DataGrid1.Columns(5).Width = 2800 '940 '4395

    DataGrid1.Columns(6).Caption = "Importe"
    DataGrid1.Columns(6).Width = 1500 '940 '4395
    DataGrid1.Columns(6).NumberFormat = "#,###,###0.00"
    DataGrid1.Columns(6).Alignment = dbgRight
    
    DataGrid1.Columns(7).Caption = "Banco"
    DataGrid1.Columns(7).Width = 600
'    DataGrid1.Columns(7).NumberFormat = "0000"
    DataGrid1.Columns(7).Alignment = dbgCenter
    
        DataGrid1.Tag = "Fijando ancho"
        txtAux(0).Width = DataGrid1.Columns(7).Width - 30
        txtAux(0).Left = DataGrid1.Left + DataGrid1.Width - txtAux(0).Width - 300
        txtAux(0).Alignment = dbgRight
        
    
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    DataGrid1.Tag = "Calculando"
    
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub CargarCombo1()
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Si, 1-No
    
    Combo1.Clear
    Combo1.AddItem "Contado"
    Combo1.ItemData(Combo1.NewIndex) = 0

    Combo1.AddItem "Efecto 19"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "Aceptada"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
    Combo1.AddItem "Recibo"
    Combo1.ItemData(Combo1.NewIndex) = 3
    
    Combo1.AddItem "Efecto 58"
    Combo1.ItemData(Combo1.NewIndex) = 4
    
End Sub


Private Function MontaSQLCarga(Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim SQL As String
    SQL = "SELECT sremes.tipofact, sremes.numserie, sremes.numfactu, sremes.ordefect, "
    SQL = SQL & "sremes.fecefect, sremes.nomclien, sremes.impefect, sremes.banremes "
    SQL = SQL & " from sremes "
    If Enlaza Then
        SQL = SQL & " WHERE numremes = " & NumRemes
    Else
        SQL = SQL & " WHERE numremes = -1"
    End If
    SQL = SQL & " ORDER BY sremes.tipofact, sremes.numserie, sremes.numfactu"
    MontaSQLCarga = SQL
End Function


Private Function Datos1Ok() As Boolean
    
    Datos1Ok = True
    If txtFec(0).Text = "" Then
        MsgBox "La Fecha de Remesa no puede estar vacia. Reintroduzca.", vbExclamation
        Datos1Ok = False
        PonerFoco txtFec(0)
        Exit Function
    End If
    
    If Combo1.ListIndex <> 1 And Combo1.ListIndex <> 2 And Combo1.ListIndex <> 4 Then
        MsgBox "El tipo de pago no es correcto. Reintroduzca.", vbExclamation
        Datos1Ok = False
        Combo1.SetFocus
        Exit Function
    End If

    ' desdes mayorees que hastas
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        If CCur(Text1(0).Text) > CCur(Text1(1).Text) Then
            MsgBox "Error entre límites. Desde no puede ser superior a hasta. ", vbExclamation
            Datos1Ok = False
            PonerFoco Text1(0)
            Exit Function
        End If
    End If
    
    If txtFec(1).Text <> "" And txtFec(2).Text <> "" Then
        If CDate(txtFec(1).Text) > CCur(txtFec(2).Text) Then
            MsgBox "Error entre límites. Desde no puede ser superior a hasta. ", vbExclamation
            Datos1Ok = False
            PonerFoco txtFec(1)
            Exit Function
        End If
    End If
End Function

Private Function DatosOk() As Boolean

    DatosOk = True
    If Text1(3).Text = "" And Text1(5).Text = "" And Text1(7).Text = "" And Text1(9).Text = "" Then
        DatosOk = False
    End If
    
End Function

Private Sub InsertaRemesa()
Dim RS As Adodb.Recordset
Dim SQL As String
Dim sql1 As String
Dim Importe As Currency
Dim OK As Boolean


            
    Set mC = New Contadores
    OK = (mC.ConseguirContador("1", 5, False) = 0)
    
    If OK Then

        Set RS = New Adodb.Recordset
        
        SQL = "select tipofact, numserie, numfactu, ordefect, fecefect,"
        SQL = SQL & "codsocio, impefect, imppagad from sefect "
        SQL = SQL & " where sefect.tipoforp = " & Combo1.ListIndex
        SQL = SQL & " and numremes is null "
        If txtFec(1).Text <> "" Then
            SQL = SQL & " and sefect.fecefect >= '" & Format(txtFec(1).Text, FormatoFecha) & "'"
        End If
        If txtFec(2).Text <> "" Then
            SQL = SQL & " and sefect.fecefect <= '" & Format(txtFec(2).Text, FormatoFecha) & "'"
        End If
        If Text1(0).Text <> "" Then
            SQL = SQL & " and (sefect.Impefect - sefect.Imppagad) >= " & TransformaComasPuntos(ImporteSinFormato(Text1(0).Text))
        Else
            SQL = SQL & " and (sefect.Impefect - sefect.Imppagad) >= 0"
        End If
        If Text1(1).Text <> "" Then
            SQL = SQL & " and (sefect.Impefect - sefect.Imppagad) <= " & TransformaComasPuntos(ImporteSinFormato(Text1(1).Text))
        End If
        
        
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        If Not RS.EOF Then RS.MoveFirst
        
        While Not RS.EOF
            Dim Nomsocio As String
            Nomsocio = ""
            Nomsocio = DevuelveDesdeBD(3, "nomlargo", "asociados", "idasoc|", RS!Codsocio & "|", "N|", 1)
            
            Importe = CCur(RS!Impefect) - CCur(RS!Imppagad)
            
            If Text1(3) <> "" Then
                If Text3(4).Text = "" Then Text3(4).Text = 0
                If CCur(Text3(4).Text) < CCur(Text1(3).Text) Then
                    sql1 = "insert into sremes(tipofact, numserie, numfactu, ordefect, banremes, fecefect,"
                    sql1 = sql1 & " impefect,  nomclien, numremes,  fecremes,  situacio) values ("
                    sql1 = sql1 & RS!tipofact & ",'" & DevNombreSQL(Trim(RS!numserie)) & "'," & Format(RS!numfactu, "0000000") & ","
                    sql1 = sql1 & RS!ordefect & "," & CCur(Text1(2).Text) & ",'" & Format(RS!FecEfect, FormatoFecha) & "',"
                    sql1 = sql1 & TransformaComasPuntos(ImporteSinFormato(CStr(Importe))) & ",'" & DevNombreSQL(Trim(Nomsocio)) & "',"
                    sql1 = sql1 & mC.Contador & ",'" & Format(txtFec(0).Text, FormatoFecha) & "',0)"
                    
                    Conn.Execute sql1
        
                    Text3(4).Text = CCur(Text3(4).Text) + Importe
                Else
                If Text1(5).Text <> "" Then
                    If Text3(9).Text = "" Then Text3(9).Text = 0
                    If CCur(Text3(9).Text) < CCur(Text1(5).Text) Then
                        sql1 = "insert into sremes(tipofact, numserie, numfactu, ordefect, banremes, fecefect,"
                        sql1 = sql1 & " impefect,  nomclien, numremes,  fecremes,  situacio) values ("
                        sql1 = sql1 & RS!tipofact & ",'" & DevNombreSQL(Trim(RS!numserie)) & "'," & Format(RS!numfactu, "0000000") & ","
                        sql1 = sql1 & RS!ordefect & "," & CCur(Text1(4).Text) & ",'" & Format(RS!FecEfect, FormatoFecha) & "',"
                        sql1 = sql1 & TransformaComasPuntos(ImporteSinFormato(CStr(Importe))) & ",'" & DevNombreSQL(Trim(Nomsocio)) & "',"
                        sql1 = sql1 & mC.Contador & ",'" & Format(txtFec(0).Text, FormatoFecha) & "',0)"
                        
                        Conn.Execute sql1
            
                        Text3(9).Text = CCur(Text3(9).Text) + Importe
                    Else
                        If Text1(7).Text <> "" Then
                            If Text3(14).Text = "" Then Text3(14).Text = 0
                            If CCur(Text3(14).Text) < CCur(Text1(7).Text) Then
                                sql1 = "insert into sremes(tipofact, numserie, numfactu, ordefect, banremes, fecefect,"
                                sql1 = sql1 & " impefect,  nomclien, numremes,  fecremes,  situacio) values ("
                                sql1 = sql1 & RS!tipofact & ",'" & DevNombreSQL(Trim(RS!numserie)) & "'," & Format(RS!numfactu, "0000000") & ","
                                sql1 = sql1 & RS!ordefect & "," & CCur(Text1(6).Text) & ",'" & Format(RS!FecEfect, FormatoFecha) & "',"
                                sql1 = sql1 & TransformaComasPuntos(ImporteSinFormato(CStr(Importe))) & ",'" & DevNombreSQL(Trim(Nomsocio)) & "',"
                                sql1 = sql1 & mC.Contador & ",'" & Format(txtFec(0).Text, FormatoFecha) & "',0)"
                                
                                Conn.Execute sql1
                    
                                Text3(14).Text = CCur(Text3(14).Text) + Importe
                            Else
                                If Text1(9).Text <> "" Then
                                    If Text3(19).Text = "" Then Text3(19).Text = 0
                                    If CCur(Text3(19).Text) < CCur(Text1(9).Text) Then
                                        sql1 = "insert into sremes(tipofact, numserie, numfactu, ordefect, banremes, fecefect,"
                                        sql1 = sql1 & " impefect,  nomclien, numremes,  fecremes,  situacio) values ("
                                        sql1 = sql1 & RS!tipofact & ",'" & DevNombreSQL(Trim(RS!numserie)) & "'," & Format(RS!numfactu, "0000000") & ","
                                        sql1 = sql1 & RS!ordefect & "," & CCur(Text1(8).Text) & ",'" & Format(RS!FecEfect, FormatoFecha) & "',"
                                        sql1 = sql1 & TransformaComasPuntos(ImporteSinFormato(CStr(Importe))) & ",'" & DevNombreSQL(Trim(Nomsocio)) & "',"
                                        sql1 = sql1 & mC.Contador & ",'" & Format(txtFec(0).Text, FormatoFecha) & "',0)"
                                        
                                        Conn.Execute sql1
                            
                                        Text3(19).Text = CCur(Text3(19).Text) + Importe
                                    End If
                                End If
                            
                            End If
                        End If
                        
                    End If
                End If
                
                End If
            End If
    
            RS.MoveNext
        Wend
    
        NumRemes = mC.Contador
    End If
    
'    Set mC = Nothing
    
End Sub

Private Function DatosBanco(codigo As String, indice As Integer) As Boolean
Dim SQL As String
Dim RS As Adodb.Recordset

    DatosBanco = False
    
    SQL = "select codbanco, codsucur, digcontr, cuentaba from sbanpr where codbanpr = "
    SQL = SQL & codigo
    
    Set RS = New Adodb.Recordset
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        RS.MoveFirst
        DatosBanco = True
        
        Select Case indice
            Case 2
                For i = 0 To 3
                    Text3(i).Text = RS.Fields(i).Value
                Next i
                'Importe de efectos restantes
                If Importe1 = 0 Then Importe1 = TotalEfectos
                If Importe1 > (Importe2 + Importe3 + Importe4) Then
                    Importe1 = TotalEfectos - Importe2 - Importe3 - Importe4
                Else
                    Importe1 = 0
                End If
                Text1(3).Text = Importe1
                
            Case 4
                For i = 5 To 8
                    Text3(i).Text = RS.Fields(i - 5).Value
                Next i
                'Importe de efectos restantes
                If Importe2 = 0 Then Importe2 = TotalEfectos
                If Importe2 > (Importe1 + Importe3 + Importe4) Then
                    Importe2 = TotalEfectos - Importe1 - Importe3 - Importe4
                Else
                    Importe2 = 0
                End If
                Text1(5).Text = Format(Importe2, "###,###,##0.00")
            
            Case 6
                For i = 10 To 13
                    Text3(i).Text = RS.Fields(i - 10).Value
                Next i
                'Importe de efectos restantes
                If Importe3 = 0 Then Importe3 = TotalEfectos
                If Importe3 > (Importe1 + Importe2 + Importe4) Then
                    Importe3 = TotalEfectos - Importe1 - Importe2 - Importe4
                Else
                    Importe3 = 0
                End If
                Text1(7).Text = Format(Importe3, "###,###,##0.00")
            
            Case 8
                For i = 15 To 18
                    Text3(i).Text = RS.Fields(i - 15).Value
                Next i
                'Importe de efectos restantes
                If Importe4 = 0 Then Importe4 = TotalEfectos
                If Importe4 > (Importe1 + Importe2 + Importe3) Then
                    Importe4 = TotalEfectos - Importe1 - Importe2 - Importe3
                Else
                    Importe4 = 0
                End If
                Text1(9).Text = Format(Importe4, "###,###,##0.00")
         End Select
        
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

Private Sub InicializarImportes()
Dim RS As Adodb.Recordset

    ' bloqueamos los campos de importes asignados
    Text1(3).Text = ""
    Text1(5).Text = ""
    Text1(7).Text = ""
    Text1(9).Text = ""
    
    Importe1 = 0
    Importe2 = 0
    Importe3 = 0
    Importe4 = 0
    TotalEfectos = 0
    
    Set RS = New Adodb.Recordset
    
    SQL = "select sum(sefect.impefect - sefect.imppagad)  from sefect "
    SQL = SQL & " where sefect.tipoforp = " & Combo1.ListIndex
    SQL = SQL & " and (numremes = 0 or numremes is null) "
    If txtFec(1).Text <> "" Then
        SQL = SQL & " and sefect.fecefect >= '" & Format(txtFec(1).Text, FormatoFecha) & "'"
    End If
    If txtFec(2).Text <> "" Then
        SQL = SQL & " and sefect.fecefect <= '" & Format(txtFec(2).Text, FormatoFecha) & "'"
    End If
    If Text1(0).Text <> "" Then
        SQL = SQL & " and (sefect.Impefect - sefect.Imppagad) >= " & TransformaComasPuntos(ImporteSinFormato(Text1(0).Text))
    End If
    If Text1(1).Text <> "" Then
        SQL = SQL & " and (sefect.Impefect - sefect.Imppagad) <= " & TransformaComasPuntos(ImporteSinFormato(Text1(1).Text))
    End If
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        RS.MoveFirst
        
        TotalEfectos = RS.Fields(0).Value
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub ModificarRegistro()
Dim SQL As String

    On Error GoTo EModificarRegistro
    

    SQL = "update sremes set banremes = " & Format(txtAux(0).Text, "0000")
    SQL = SQL & " where tipofact = " & Adodc1.Recordset.Fields(0).Value
    SQL = SQL & " and numserie = '" & Adodc1.Recordset.Fields(1).Value & "' "
    SQL = SQL & " and numfactu = " & Adodc1.Recordset.Fields(2).Value
    SQL = SQL & " and ordefect = " & Adodc1.Recordset.Fields(3).Value
    
    Conn.Execute SQL

EModificarRegistro:
    If Err.Number <> 0 Then
        MuestraError 0, "Error en la modificacion de registro."
    Else
        RecalculoImportes
    End If
End Sub

Private Function DatosBanOk() As Boolean
    DatosBanOk = False
    If (txtAux(0).Text = Text1(2).Text) Or _
       (txtAux(0).Text = Text1(4).Text) Or _
       (txtAux(0).Text = Text1(6).Text) Or _
       (txtAux(0).Text = Text1(8).Text) Then
        DatosBanOk = True
    End If

End Function

Private Sub RecalculoImportes()
Dim RS As Adodb.Recordset

    Text3(4).Text = ""
    Text3(9).Text = ""
    Text3(14).Text = ""
    Text3(19).Text = ""
    
    
    Set RS = New Adodb.Recordset
    
    If Text1(2).Text <> "" Then
        SQL = "select sum(sremes.impefect) from sremes "
        SQL = SQL & " where numremes = " & NumRemes
        SQL = SQL & " and banremes =  " & Text1(2).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Text3(4).Text = ""
        If Not RS.EOF Then
            RS.MoveFirst
            If Not IsNull(RS.Fields(0)) Then Text3(4).Text = RS.Fields(0).Value
        End If
        
        RS.Close
    End If
    If Text1(4).Text <> "" Then
        SQL = "select sum(sremes.impefect) from sremes "
        SQL = SQL & " where numremes = " & NumRemes
        SQL = SQL & " and banremes =  " & Text1(4).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Text3(9).Text = ""
        If Not RS.EOF Then
            RS.MoveFirst
            If Not IsNull(RS.Fields(0)) Then Text3(9).Text = RS.Fields(0).Value
        End If
        
        RS.Close
    End If
    If Text1(6).Text <> "" Then
        SQL = "select sum(sremes.impefect) from sremes "
        SQL = SQL & " where numremes = " & NumRemes
        SQL = SQL & " and banremes =  " & Text1(6).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Text3(14).Text = ""
        If Not RS.EOF Then
            RS.MoveFirst
            If Not IsNull(RS.Fields(0)) Then Text3(14).Text = RS.Fields(0).Value
        End If
        
        RS.Close
    End If
    If Text1(8).Text <> "" Then
        SQL = "select sum(sremes.impefect) from sremes "
        SQL = SQL & " where numremes = " & NumRemes
        SQL = SQL & " and banremes =  " & Text1(8).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
        Text3(19).Text = ""
        If Not RS.EOF Then
            RS.MoveFirst
            If Not IsNull(RS.Fields(0)) Then Text3(19).Text = RS.Fields(0).Value
        End If
        
        RS.Close
    End If
        
    Set RS = Nothing

End Sub

Private Sub ActualizaCartera()
    Adodc1.Recordset.MoveFirst
    While Not Adodc1.Recordset.EOF
        SQL = "update sefect set numremes = " & NumRemes & ", fecremes = '"
        SQL = SQL & Format(txtFec(0).Text, FormatoFecha) & "', banremes = "
        SQL = SQL & Adodc1.Recordset.Fields(7).Value
        SQL = SQL & " where tipofact = " & Adodc1.Recordset.Fields(0).Value & " and "
        SQL = SQL & " numserie = '" & Adodc1.Recordset.Fields(1).Value & "' and "
        SQL = SQL & " numfactu = " & Adodc1.Recordset.Fields(2).Value & " and "
        SQL = SQL & " ordefect =  " & Adodc1.Recordset.Fields(3).Value
         
        Conn.Execute SQL
        
        Adodc1.Recordset.MoveNext
    Wend
End Sub

Private Sub BloqueoCampos(Index As Integer)
Dim i As Integer
    If Index < 2 Or Index > 9 Then Exit Sub

    If Text1(Index) = "" Then
        For i = Index + 1 To 9
            Text1(i).Enabled = False
        Next i
    Else
        Text1(Index + 1).Enabled = True
        If Index < 8 Then Text1(Index + 2).Enabled = True
    End If
    
End Sub

