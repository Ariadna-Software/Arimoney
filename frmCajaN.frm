VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCajaN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducción de caja"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmCajaN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   4560
      TabIndex        =   7
      Top             =   6240
      Width           =   885
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   43
      Text            =   "Text3"
      Top             =   810
      Width           =   3855
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6240
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Tag             =   "CODUSU|N|N|0||scacaja|codusu||S|"
      Text            =   "Text1"
      Top             =   810
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha |F|N|||scacaja|feccaja|dd/mm/yyyy|S|"
      Text            =   "commor"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   19
      Top             =   6240
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9600
      TabIndex        =   18
      Top             =   7440
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   1080
      TabIndex        =   33
      Top             =   6240
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3420
      TabIndex        =   6
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   5760
      TabIndex        =   8
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   6480
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1515
      Left            =   7800
      TabIndex        =   24
      Top             =   360
      Width           =   2895
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "COBROS"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "PAGOS"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   660
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2520
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   120
      TabIndex        =   20
      Top             =   7320
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      TabIndex        =   17
      Top             =   7440
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCajaN.frx":030A
      Height          =   4455
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7858
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
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
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "CIERRE CAJA"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7680
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   495
      Left            =   5520
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
   Begin VB.Frame framelineas2 
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   6360
      Width           =   10515
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   4680
         TabIndex        =   32
         Text            =   "Text3"
         Top             =   360
         Width           =   5175
      End
      Begin VB.Frame FrameCombo 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   4335
         Begin VB.CommandButton cmdTrer 
            Caption         =   "+"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   10
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   11
            Left            =   2520
            TabIndex        =   15
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   12
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   16
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label2 
            Caption         =   "proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   57
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Factura "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   56
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Numero"
            Height          =   195
            Index           =   11
            Left            =   1440
            TabIndex        =   55
            Top             =   0
            Width           =   570
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   3120
            Picture         =   "frmCajaN.frx":031F
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   10
            Left            =   2520
            TabIndex        =   54
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label1 
            Caption         =   "Vto"
            Height          =   195
            Index           =   9
            Left            =   3600
            TabIndex        =   53
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.Frame FrameCombo 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   4335
         Begin VB.CommandButton cmdTrer 
            Caption         =   "+"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   3
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   9
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   13
            Top             =   240
            Width           =   405
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   8
            Left            =   2520
            TabIndex        =   12
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   7
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   320
            Index           =   6
            Left            =   840
            MaxLength       =   3
            TabIndex        =   10
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label2 
            Caption         =   "cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   58
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Vto"
            Height          =   195
            Index           =   7
            Left            =   3600
            TabIndex        =   51
            Top             =   0
            Width           =   450
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   6
            Left            =   2520
            TabIndex        =   50
            Top             =   0
            Width           =   570
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   3120
            Picture         =   "frmCajaN.frx":03AA
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Numero"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   49
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label1 
            Caption         =   "Serie"
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   48
            Top             =   0
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Factura "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   47
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Factura "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   8
         Left            =   360
         TabIndex        =   59
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Forma pago"
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
         Index           =   1
         Left            =   4680
         TabIndex        =   45
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   5880
         Picture         =   "frmCajaN.frx":0435
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame frameextras2 
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   6360
      Width           =   10575
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   4
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text3"
         Top             =   420
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   40
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Forma pago"
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
         Index           =   4
         Left            =   5040
         TabIndex        =   39
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "C A J A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4080
      TabIndex        =   60
      Top             =   7320
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Forma pago"
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
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmCajaN.frx":6C87
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   42
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   41
      Top             =   1200
      Width           =   450
   End
   Begin VB.Menu mnOpcionesAsiPre 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
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
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "Lineas"
         Shortcut        =   ^L
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
Attribute VB_Name = "frmCajaN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public UsuarioCajaPredeterminada As Boolean
        
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmFormaPago
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1

Private WithEvents frmCob As frmCobros
Attribute frmCob.VB_VarHelpID = -1
Private WithEvents frmpag As frmPagoPro
Attribute frmpag.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busquedaa
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'/////////////////////////////////////
'////////////////////////////   //////
'//////////////////////////////////
'   Nuevo modo --> Modificando lineas
'  5.- Modificando lineas

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private SQL As String
Dim I As Integer



Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas


'Para pasar de lineas a cabeceras
Dim Linliapu As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar


Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean

Dim CargandoGrid As Byte   '0 no hace nada. 1 .- Es el primero de los dos que hace al poner adtos sobre el grid

Dim PosicionGrid As Integer

Private CadenaAmpliacion As String
Private DatosPartidasPendientesAplicacion As String

Private MaxImporteFactura As Currency
Private ImporteYaCobrado As Currency 'En la factura

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If CargandoGrid > 0 Then
        If CargandoGrid = 1 Then
            'La primera vez:
            CargandoGrid = 2
        Else
            'AQUI PONGO LOS DATOS en los txt correspondientes
            PonerDatosFactura
            
        End If
    End If
End Sub





Private Sub cmbTipo_Click()
    'Cuando pierda el foco entonces pondre un frame u otr
    If Modo = 5 And ModificandoLineas = 1 Then
        'INSERTANDO LINEAS
        PonerFrameCombo cmbTipo.ListIndex
        
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmbTipo_LostFocus()
    If Modo = 5 And ModificandoLineas = 1 Then
        Select Case cmbTipo.ListIndex = 2
        Case 0, 1
            If cmbTipo.ListIndex >= 0 Then Ponerfoco Me.cmdTrer(cmbTipo.ListIndex)
    
        Case 2
            If txtaux(0).Text = "" And txtaux(1).Text = "" Then
                txtaux(0).Text = RecuperaValor(DatosPartidasPendientesAplicacion, 1)
                txtaux(1).Text = RecuperaValor(DatosPartidasPendientesAplicacion, 2)
            End If
        End Select
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim I As Integer
    Dim Limp As Boolean
    Dim B As Boolean
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            
            
            
            If I > 2 Then
                cad = "Fecha fuera de ejercicio contable." & vbCrLf & " Continuar?"
                If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then
                    B = False
                Else
                    B = True
                End If
            Else
                    cmdCancelar.Caption = "Cancelar"
    
                    B = InsertarDesdeForm(Me)
            End If
            
            If B Then

                        
                    'Ponemos la cadena consulta
                    If SituarData1(True) Then
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        cmdCancelar.Caption = "Cabecera"
                        
                        ModificandoLineas = 0
                        AnyadirLinea True
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: " & Me.Name & " cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                        Exit Sub
                    End If

            End If  'de B
        End If   'datosok
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos modificar
                'PreparaBloquear
                Limp = Modificar
                'TerminaBloquear
                If Limp Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1(False) Then
                        lblIndicador.Caption = ""
                        PonerModo 2
                    Else
                        PonerModo 0
                    End If
                Else
                    PonerCampos
                End If
            End If
            
    Case 5
        cad = AuxOK
        If cad <> "" Then
            If cad <> "NO" Then MsgBox cad, vbExclamation
        Else
            'Insertaremos, o modificaremos
            If InsertarModificar Then
                'Reestablecemos los campos
                'y ponemos el grid
                cmdAceptar.Visible = False
                DataGrid1.AllowAddNew = False
                
                If Not Adodc1.Recordset.EOF Then PosicionGrid = DataGrid1.FirstRow
                CargaGrid True
                Limp = True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    'Si ha puesto contrapartida borramos
                    
                    txtaux(8).Text = ""
                    Text3(2).Text = ""
                    If Limp Then
                        For I = 4 To 5
                            Text3(I).Text = ""
                        Next I
                        For I = 0 To txtaux.Count - 1
                            txtaux(I).Text = ""
                        Next I
                    End If
                    ModificandoLineas = 0
                    cmdAceptar.Visible = True
                    cmdCancelar.Caption = "C&abecera"
                    AnyadirLinea False

                Else
                    ModificandoLineas = 0
                    
                    'Intentamos poner el grid donde toca
                    PonerLineaModificadaSeleccionada
                    CamposAux False, 0, False
                    cmdCancelar.Caption = "Cabecera"
                    lblIndicador.Caption = ""
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

Private Sub PonerLineaModificadaSeleccionada()
    On Error GoTo E1
   ' While Not Adodc1.Recordset.EOF
   '     If CStr(Adodc1.Recordset.Fields(1)) = CStr(Linliapu) Then Exit Sub
   '     Adodc1.Recordset.MoveNext
   ' Wend
   
    
   Adodc1.Recordset.Find "numlinea =" & Linliapu
 
   
   If Adodc1.Recordset.RecordCount - Adodc1.Recordset.AbsolutePosition < DataGrid1.VisibleRows Then
        'Estoy en la utlimo trozo. No habra scroll
   Else
        I = PosicionGrid - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
    End If
    Exit Sub
E1:
    Err.Clear
End Sub



Private Sub cmdAux_Click(Index As Integer)
    
    

        Set frmCCtas = New frmColCtas

        frmCCtas.Busqueda = CadenaAmpliacion
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.ConfigurarBalances = 7
        CadenaAmpliacion = ""
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        If CadenaAmpliacion <> "" Then
            txtaux(0).Text = RecuperaValor(CadenaAmpliacion, 1)
            txtaux(1).Text = RecuperaValor(CadenaAmpliacion, 2)
            CadenaAmpliacion = ""
        End If
    
    'txtAux_LostFocus Index
    If txtaux(0).Text <> "" Then Ponerfoco txtaux(2)
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3
        LimpiarCampos
        PonerModo 0
    Case 4
        lblIndicador.Caption = ""
        PonerModo 2
        PonerCampos
        
        
    Case 5
         lblIndicador.Caption = ""
        CamposAux False, 0, False
        frameextras2.Visible = True
        framelineas2.Visible = False
 

        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
'            If Adodc1.Recordset.EOF Then
'                SQL = "El asiento no tiene lineas. Desea salir igualmente?"
'                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
'            Else
'                'Si el asiento esta descuadrado hbar que dar una notificacion
'                If Text2(2).Text <> "" Then
'                    SQL = "El asiento esta descuadrado. Seguro que desea salir de la edición de lineas de asiento ?"
'                    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
'                Else
'                    'Si asiento cuadrado y actualizar automaticamente
'                    'lanzamos actualizacion
'
'                End If
'            End If
           lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If
            frameextras2.Visible = Not Adodc1.Recordset.EOF
            cmdAceptar.Visible = False
            cmdCancelar.Caption = "Cabeceras"
            ModificandoLineas = 0
        End If
    End Select
End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim SQL As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        SQL = "Select * from scacaja WHERE codusu =" & Text1(4).Text
        SQL = SQL & " AND feccaja='" & Format(Text1(1).Text, FormatoFecha) & "' "
        Data1.RecordSource = SQL
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!codusu) = Text1(4).Text Then
                If Format(CStr(.Fields!feccaja), "dd/mm/yyyy") = Text1(1).Text Then
                    SituarData1 = True
                    Exit Function
                End If
            End If
            .MoveNext
        Wend
    End With
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
    'CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    'PonerCadenaBusqueda True
    
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3

    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    '###A mano
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Ponerfoco Text1(1)
    
    
    'El usuario es el k es
    I = vUsu.Codigo Mod 100
    Text1(4).Text = I
    Text1(3).Text = vUsu.Nombre
    
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Ponerfoco Text1(4)
        Text1(4).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Ponerfoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia "codusu = " & (vUsu.Codigo Mod 100)
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        If Not UsuarioCajaPredeterminada Then CadenaConsulta = CadenaConsulta & " WHERE codusu = " & (vUsu.Codigo Mod 100)
        CadenaConsulta = CadenaConsulta & Ordenacion
        PonerCadenaBusqueda False
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
If Data1.Recordset.EOF Then Exit Sub
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
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    
    
    
    
    'Comprobamos que la fecha es de ejerccio actual
    If FechaCorrecta2(CDate(Text1(1).Text), False) = 2 Then
        MsgBox "Fecha fuera ejercicios o de ambito.", vbExclamation
        Exit Sub
    End If
    
    
   
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdCancelar.Caption = "Cancelar"
    cmdAceptar.Caption = "&Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano

    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    Ponerfoco Text1(0)
End Sub




Private Sub BotonEliminar()
        On Error GoTo Error2
        
    SQL = vbCrLf & String(30, "=") & vbCrLf
    SQL = SQL & "¿Desea eliminar la caja:" & vbCrLf & "Usuario: " & Text1(4).Text & " - " & Text1(3).Text & _
         vbCrLf & "Fecha: " & Text1(1).Text & SQL
         
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    NumRegElim = 0
    If Not Adodc1.Recordset.EOF Then
        NumRegElim = DBLet(Adodc1.Recordset.RecordCount, "N")
    End If
    If NumRegElim > 0 Then
        SQL = "Seguro que quiere eliminar la caja:" & vbCrLf & "Usuario: " & Text1(4).Text & " - " & Text1(3).Text
        SQL = SQL & vbCrLf & "Fecha: " & Text1(1).Text
        SQL = SQL & vbCrLf & "Lineas: " & NumRegElim
        SQL = SQL & vbCrLf & vbCrLf & vbCrLf & "¿ CONTINUAR ? "
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    End If
    
    If Not Eliminar Then Exit Sub


    
    Screen.MousePointer = vbHourglass
    NumRegElim = Data1.Recordset.AbsolutePosition
    DataGrid1.Enabled = False
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim - 1
            End If
            PonerCampos
            DataGrid1.Enabled = True
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If

Error2:
        Screen.MousePointer = vbDefault
        
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdClien_Click()

End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

'If Data1.Recordset.EOF Then
'    MsgBox "Ningún registro devuelto.", vbExclamation
'    Exit Sub
'End If
'
'Cad = ""
'i = 0
'Do
'    j = i + 1
'    i = InStr(j, DatosADevolverBusqueda, "|")
'    If i > 0 Then
'        AUX = Mid(DatosADevolverBusqueda, j, i - j)
'        j = Val(AUX)
'        Cad = Cad & Text1(j).Text & "|"
'    End If
'Loop Until i = 0
'RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub







Private Sub cmdTrer_Click(Index As Integer)
    
    HaDevueltoDatos = False
    
    If Index = 0 Then
        'Cobros
        Set frmCob = New frmCobros
        frmCob.DatosADevolverBusqueda = "SI"
        frmCob.Show vbModal
        Set frmCob = Nothing
        
    Else
        'pagos
        Set frmpag = New frmPagoPro
        frmpag.DatosADevolverBusqueda = "SI"
        frmpag.Show vbModal
        Set frmpag = Nothing
    End If
    
    If HaDevueltoDatos Then
        HabilitarCamposFacturas False
        Ponerfoco cmdAceptar
    End If
    
End Sub

Private Sub Form_Activate()

  
    If PrimeraVez Then

        PrimeraVez = False
      
        Modo = 0
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codusu = -1"
        Data1.ConnectionString = Conn
        Data1.RecordSource = CadenaConsulta
        Data1.Refresh
        PonerModo CInt(Modo)
        CargaGrid (Modo = 2)
        Toolbar1.Enabled = True
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .Enabled = False
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 17
        .Buttons(13).Image = 16
        .Buttons(14).Image = 15
        .Buttons(16).Image = 6
        .Buttons(17).Image = 7
        .Buttons(18).Image = 8
        .Buttons(19).Image = 9
    End With
    
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
       ' Me.Width = 12000
       ' Me.Height = Screen.Height
    End If
    Me.Height = 8605
    'Los campos auxiliares
    CamposAux False, 0, True
    
    'Si no es analitica no mostramos el label, texto ni IMAGEN
    'Text3(2).Visible = vParam.autocoste
    'Label2(2).Visible = vParam.autocoste
    'Image1(2).Visible = vParam.autocoste
    
    CargaCombo
    PonerDatosPartidasPendientesAplicacion
    
    '## A mano
    NombreTabla = "scacaja"
    Ordenacion = " ORDER BY feccaja,codusu"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
'    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.Password = vUsu.Passwd
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login

    
    'Maxima longitud cuentas
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    
    'Bloqueo de tabla, cursor type
'    Data1.CursorType = adOpenDynamic
'    Data1.LockType = adLockPessimistic
    'CadAncho = False
    PulsadoSalir = False
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
End Sub


'Private Sub Form_Resize()
'If Me.WindowState <> 0 Then Exit Sub
'If Me.Width < 11610 Then Me.Width = 11610
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim B As Boolean
    
    If Modo > 2 Then
        B = True
    Else

    End If
    If B Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        

        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE "
        CadenaConsulta = CadenaConsulta & CadB
        PonerCadenaBusqueda False
        Screen.MousePointer = vbDefault
    End If

End Sub

'Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
''Cuentas
'If cmdAux(0).Tag = 0 Then
'    'Cuenta normal
'    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
'    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
'
'
'Else
'    'contrapartida
'    txtaux(3).Text = RecuperaValor(CadenaSeleccion, 1)
'    Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
'End If
'End Sub






'
'Private Sub frmF_Selec(vFecha As Date)
'Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub



Private Sub frmC_Selec(vFecha As Date)
    If I >= 1000 Then
        Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
    Else
        txtaux(I).Text = Format(vFecha, "dd/mm/yyyy")
    End If
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    CadenaAmpliacion = CadenaSeleccion
End Sub

Private Sub frmCob_DatoSeleccionado(CadenaSeleccion As String)

    
    'Datos factura
    SQL = ""
    For I = 1 To 4
        
        txtaux(5 + I).Text = RecuperaValor(CadenaSeleccion, I)
        SQL = SQL & " " & txtaux(5 + I).Text & " "
    Next I
    
    'Ampliacion
    txtaux(3).Text = "COBRO FACT: " & Trim(SQL)
    
    'Codmacta nommacta
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 5)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 6)
    'Codforpa nomforpa
    txtaux(2).Text = RecuperaValor(CadenaSeleccion, 7)
    Text3(2).Text = RecuperaValor(CadenaSeleccion, 8)
    'Importe
    SQL = RecuperaValor(CadenaSeleccion, 9)
    MaxImporteFactura = ImporteFormateado(SQL)
    
    
    txtaux(4).Text = CStr(MaxImporteFactura)
    txtaux(5).Text = ""
    
    SQL = RecuperaValor(CadenaSeleccion, 10)
    If SQL = "" Then SQL = 0
    ImporteYaCobrado = ImporteFormateado(SQL)
    
    HaDevueltoDatos = True
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(2).Text = RecuperaValor(CadenaSeleccion, 1)
    'Text3(2).Text = RecuperaValor(CadenaSeleccion, 2) & " (" & RecuperaValor(CadenaSeleccion, 3) & ")"
    Text3(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmpag_DatoSeleccionado(CadenaSeleccion As String)

    'Datos factura
    SQL = ""
    For I = 1 To 3
        
        txtaux(9 + I).Text = RecuperaValor(CadenaSeleccion, I)
        SQL = SQL & " " & txtaux(9 + I).Text & " "
    Next I
    
    'Ampliacion
    txtaux(3).Text = "PAGO FACT: " & Trim(SQL)
    
    'Codmacta nommacta
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 4)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 5)
    'Codforpa nomforpa
    txtaux(2).Text = RecuperaValor(CadenaSeleccion, 6)
    Text3(2).Text = RecuperaValor(CadenaSeleccion, 7)
    'Importe
    SQL = RecuperaValor(CadenaSeleccion, 8)
    MaxImporteFactura = ImporteFormateado(SQL)
    
    txtaux(5).Text = CStr(MaxImporteFactura)
    txtaux(4).Text = ""
    
    
    SQL = RecuperaValor(CadenaSeleccion, 9)
    If SQL = "" Then SQL = 0
    ImporteYaCobrado = ImporteFormateado(SQL)
    
    HaDevueltoDatos = True
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0, 1
    If Index = 0 Then
        I = 8
    Else
        I = 11
    End If
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtaux(I).Text <> "" Then frmC.Fecha = CDate(txtaux(I).Text)
    frmC.Show vbModal
    Set frmC = Nothing
Case 2
    'FORPA
    Set frmF = New frmFormaPago
    frmF.DatosADevolverBusqueda = "0|1|"
    frmF.Show vbModal
    Set frmF = Nothing
    If txtaux(2).Text <> "" Then Ponerfoco txtaux(3)
    
End Select
End Sub

Private Sub imgppal_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0
        'FECHA
        I = 1000
        Set frmC = New frmCal
        frmC.Fecha = Now
        If Text1(1).Text <> "" Then frmC.Fecha = CDate(Text1(1).Text)
        frmC.Show vbModal
        Set frmC = Nothing
'    Case 1
'        'Tipos diario
'        Set frmDi = New frmTiposDiario
'        frmDi.DatosADevolverBusqueda = "0"
'        frmDi.Show vbModal
'        Set frmDi = Nothing
'    Case 2
'        'ASiento predefinido
'        If Modo = 3 Then
'            'Solo si es nuevo
'            Set frmPre = New frmAsiPre
'            frmPre.DatosADevolverBusqueda = "0"
'            frmPre.Show vbModal
'            Set frmPre = Nothing
'        End If
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    'BotonEliminar False
    HacerToolBar 8
End Sub

Private Sub mnLineas_Click()
Dim B As Button
    Set B = Toolbar1.Buttons(10)
    Toolbar1_ButtonClick B
    Set B = Nothing
End Sub

Private Sub mnModificar_Click()
    'BotonModificar
    HacerToolBar 7
End Sub

Private Sub mnNuevo_Click()
    'BotonAnyadir
    HacerToolBar 6
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    If Modo = 5 Then Exit Sub
    PulsadoSalir = True
    Screen.MousePointer = vbHourglass
    DataGrid1.Enabled = False
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


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
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
Dim RC As Byte
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    'Si estamos insertando o modificando o buscando
    If Modo = 3 Or Modo = 4 Then
        If Text1(Index).Text = "" Then
            If Index = 0 Then
                'Text4.Text = ""
            Else
                
            End If
            Exit Sub
        End If
        Select Case Index
        Case 0

        Case 1
            SQL = ""
            If Not EsFechaOK(Text1(1)) Then
                MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
                SQL = "mal"
            Else
                RC = FechaCorrecta2(CDate(Text1(1).Text), True)
                'Text1(1).Text = Format(Text1(1).Text, "dd/mm/yyyy")
                SQL = ""
                If RC > 1 Then SQL = "MAL"
                    
            End If
            If SQL <> "" Then
                Text1(1).Text = ""
                Ponerfoco Text1(1)
            End If
            
        Case 2

        End Select
    End If
End Sub

Private Sub HacerBusqueda()
    Dim cad As String
    Dim CadB As String
    CadB = ObtenerBusqueda(Me)
    
    If Not UsuarioCajaPredeterminada Then
        If CadB <> "" Then CadB = " AND " & CadB
        CadB = " codusu = " & (vUsu.Codigo Mod 100) & CadB
    End If
    
    If chkVistaPrevia = 1 Then
            MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda False
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim cad As String
        'Llamamos a al form
        '##A mano
        cad = ""
        cad = cad & ParaGrid(Text1(4), 20, "Usuario:")
        cad = cad & ParaGrid(Text1(1), 30, "Fecha")
        'cad = cad & ParaGrid(Text1(0), 15, "Nº Diario")
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Caja"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                'If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
               ' Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda(Insertando As Boolean)
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Insertando Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla de Apuntes", vbInformation
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
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    Text1(3).Text = DevuelveNombreUsuario(CInt(Data1.Recordset!codusu))
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True

    frameextras2.Visible = Not Adodc1.Recordset.EOF

    If Modo = 2 Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean


    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
'        For i = 0 To Text1.Count - 1
'            Text1(i).BackColor = vbWhite
'            'Text1(0).BackColor = &H80000018
'        Next i
'        'chkVistaPrevia.Visible = False
        'Reestablecemos el color del nuº asien
        Text1(4).BackColor = &HFEF7E4
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nueva caja"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar caja"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar caja"
    End If
    

        
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea de caja"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea de caja"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea de caja"
    End If
    PonerOpcionesMenuGeneral Me
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    frameextras2.Visible = B
    If B Then framelineas2.Visible = False
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B
    Toolbar1.Buttons(11).Enabled = B
    If Not B Then frameextras2.Visible = False
        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.Visible = B Or Modo = 1
    'PRueba###
    


    '
    B = B Or (Modo = 5 And ModificandoLineas = 0)
    'mnOpcionesAsiPre.Enabled = Not B
    'Los buscar ver todos y salir estaran en lo que seria el enabled
    mnBuscar.Enabled = Not B
    mnVerTodos.Enabled = Not B
    mnSalir.Enabled = Not B
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
   
   
        'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5

    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = (Modo = 2)
'    Else
'        cmdRegresar.Visible = False
'    End If
    
    '
    Text1(4).Enabled = (Modo = 1)

    B = (Modo = 3) Or (Modo = 4) Or (Modo = 1)

    Text1(1).Enabled = B

    'El text
    B = (Modo = 2) Or (Modo = 5)
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B


   
   
    If Modo <= 2 Then
         Me.cmdAceptar.Caption = "Aceptar"
         Me.cmdCancelar.Caption = "Cancelar"
    End If
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    
    B = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = B
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    
    For I = 6 To 11
        If I <> 9 Then Me.Toolbar1.Buttons(I).Enabled = Me.Toolbar1.Buttons(I).Enabled And vUsu.Nivel < 3
    Next I
    
    
    Me.mnNuevo.Enabled = Me.Toolbar1.Buttons(6).Enabled
    Me.mnEliminar.Enabled = Me.Toolbar1.Buttons(7).Enabled
    Me.mnModificar.Enabled = Me.Toolbar1.Buttons(8).Enabled
    Me.mnLineas.Enabled = Me.Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(13).Enabled = False
End Sub


Private Function DatosOk() As Boolean
    Dim Rs As ADODB.Recordset
    Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    
    If FechaCorrecta2(CDate(Text1(1).Text), True) > 1 Then B = False
        
    
    
    
    If Modo = 4 Then
        If CDate(Text1(1).Text) <> Data1.Recordset!feccaja Then
        'MODIFICAR. Compruebo que no exista otra caja para esta fecha y este usuario
            Set miRsAux = New ADODB.Recordset
            SQL = "Select count(*) from scacaja where codusu = " & Text1(4).Text & " AND feccaja = '" & Format(Text1(1).Text, FormatoFecha) & "'"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If DBLet(miRsAux.Fields(0), "N") > 0 Then
                    MsgBox "Ya existe la caja para la fecha: " & Text1(1).Text & " y el usuario: " & Text1(4).Text, vbExclamation
                    B = False
                End If
            End If
            miRsAux.Close
            Set miRsAux = Nothing
        End If
    End If
    DatosOk = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub HacerToolBar(Boton As Integer)

    If Boton > 6 And Boton <= 11 Then
         If Not (Data1.Recordset Is Nothing) Then
            If Not Data1.Recordset.EOF Then
                If (vUsu.Codigo Mod 100) <> Val(CStr(Data1.Recordset!codusu)) Then
                    'No tiene permisoso
                    MsgBox "No puede realizar esta accion sobre la caja", vbExclamation
                    Exit Sub
                End If
            End If
        End If

    End If

    Select Case Boton
    Case 1
        BotonBuscar
    Case 2
        BotonVerTodos
    Case 6
        If Modo <> 5 Then
            BotonAnyadir
        Else
            'AÑADIR linea factura
            AnyadirLinea True
        End If
    Case 7
        If Modo <> 5 Then
            'Intentamos bloquear la cuenta
            If Data1.Recordset Is Nothing Then Exit Sub
            If Data1.Recordset.EOF Then Exit Sub

            BotonModificar
        Else
            'MODIFICAR linea factura
            ModificarLinea
        End If
    Case 8
        If Modo <> 5 Then
            BotonEliminar
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
        'If RecodsetVacio Then Exit Sub

        'Nuevo Modo
        PonerModo 5
        'Fuerzo que se vean las lineas
        frameextras2.Visible = True
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
    Case 11
        'ACtualizar CAJA
        If Data1.Recordset.EOF Then
            MsgBox "Ningúna caja para cerrar.", vbExclamation
            Exit Sub
        End If
        If Adodc1 Is Nothing Then Exit Sub
        If Adodc1.Recordset.EOF Then
            MsgBox "No hay lineas insertadas para esta caja", vbExclamation
            Exit Sub
        End If
        
        
        
        SQL = "Va a cerrar la caja: " & vbCrLf & "Dia : " & Text1(1).Text & vbCrLf & _
            "Numero lineas : " & Adodc1.Recordset.RecordCount & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then ContabilizarCaja
        
        
    Case 13
        'Imprimir asientos


    
    Case 14
        'SALIR
        If Modo < 3 Then mnSalir_Click
    Case 16 To 19
        Desplazamiento (Boton - 16)
    Case Else
    
    End Select
End Sub







Private Sub DespalzamientoVisible(Bol As Boolean)
    For I = 16 To 19
        Toolbar1.Buttons(I).Enabled = Bol
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub



Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    
    CargandoGrid = Abs(Enlaza)
    
        
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
    
    'Claves lineas asientos predefinidos
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Visible = False
    
    'Cuenta
    DataGrid1.Columns(2).Caption = "Tipo"
    DataGrid1.Columns(2).Width = 855
    
    DataGrid1.Columns(3).Caption = "Cuenta"
    DataGrid1.Columns(3).Width = 1200


    DataGrid1.Columns(4).Caption = "Descripcion"
    DataGrid1.Columns(4).Width = 2600

    DataGrid1.Columns(5).Caption = "F.P."
    DataGrid1.Columns(5).Width = 600
    
    DataGrid1.Columns(6).Caption = "Ampliación"
    DataGrid1.Columns(6).Width = 2400
    
    
    DataGrid1.Columns(7).Caption = "Cobros"
    DataGrid1.Columns(7).NumberFormat = FormatoImporte
    DataGrid1.Columns(7).Width = 1154
    DataGrid1.Columns(7).Alignment = dbgRight
            
    DataGrid1.Columns(8).Caption = "Pagos"
    DataGrid1.Columns(8).NumberFormat = FormatoImporte
    DataGrid1.Columns(8).Width = 1154
    DataGrid1.Columns(8).Alignment = dbgRight
            
    Dim II As Integer
    For II = 9 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(II).Visible = False
    Next II
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        anc = 323
        
        'Primero el combo
        Me.cmbTipo.Left = DataGrid1.Left + 330
        
        
        
        txtaux(0).Left = DataGrid1.Columns(3).Left + 150
        txtaux(0).Width = DataGrid1.Columns(3).Width - 60
        
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(4).Left + 90
                
        txtaux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 6
        txtaux(1).Width = DataGrid1.Columns(4).Width - 180
    
        txtaux(2).Left = DataGrid1.Columns(5).Left + 150
        txtaux(2).Width = DataGrid1.Columns(5).Width - 30
    
        txtaux(3).Left = DataGrid1.Columns(6).Left + 150
        txtaux(3).Width = DataGrid1.Columns(6).Width - 45

        
        'Concepto
        txtaux(4).Left = DataGrid1.Columns(7).Left + 150
        txtaux(4).Width = DataGrid1.Columns(7).Width - 30
        
        txtaux(5).Left = DataGrid1.Columns(8).Left + 150
        txtaux(5).Width = DataGrid1.Columns(8).Width - 30

        CadAncho = True
    End If
        
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    DataGrid1.Tag = "Calculando"
    'Obtenemos las sumas
    ObtenerSumas
    
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub ObtenerSumas()
    Dim Deb As Currency
    Dim hab As Currency
    Dim Rs As ADODB.Recordset
    
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    
    If Data1.Recordset Is Nothing Then Exit Sub
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Adodc1.Recordset Is Nothing Then Exit Sub
    
    If Adodc1.Recordset.EOF Then Exit Sub
    
    
    Set Rs = New ADODB.Recordset
    
    SQL = "SELECT Sum(slicaja.importeD) AS SumaDetimporteD, Sum(slicaja.importeH) AS SumaDetimporteH"
    SQL = SQL & " ,slicaja.codusu,slicaja.feccaja"
    SQL = SQL & " From slicaja GROUP BY slicaja.codusu,slicaja.feccaja "
    SQL = SQL & " HAVING (((slicaja.codusu)=" & Data1.Recordset!codusu
    SQL = SQL & ") AND ((slicaja.feccaja)='" & Format(Data1.Recordset!feccaja, FormatoFecha)
    SQL = SQL & "'));"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Deb = 0
    hab = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Deb = Rs.Fields(0)
        If Not IsNull(Rs.Fields(1)) Then hab = Rs.Fields(1)
    End If
    Rs.Close
    Set Rs = Nothing
    Text2(0).Text = Format(Deb, FormatoImporte): Text2(1).Text = Format(hab, FormatoImporte)
    'Metemos en DEB el total
    Deb = Deb - hab
    If Deb < 0 Then
        Text2(2).ForeColor = vbRed
        Else
        Text2(2).ForeColor = vbBlack
    End If
    If Deb <> 0 Then Text2(2).Text = Format(Deb, FormatoImporte)
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
    SQL = "select codusu, slicaja.tipomovi, stipcaja.siglas"
    SQL = SQL & ",slicaja.codmacta, Nommacta, slicaja.codforpa, ampliaci, ImporteD, ImporteH"
    'los campos siguientes NO, repito NO, van en el grid
    SQL = SQL & " ,feccaja,numlinea,numserie, numfaccl, fecfaccl, numfacpr,fecfacpr,numvenci,nomforpa,descformapago"
    SQL = SQL & " From slicaja, stipcaja, cuentas,sforpa,stipoformapago"
    SQL = SQL & " Where slicaja.tipomovi = stipcaja.tipomovi And slicaja.codmacta = cuentas.codmacta AND "
    SQL = SQL & " slicaja.codforpa = sforpa.codforpa AND sforpa.tipforpa = stipoformapago.tipoformapago and"
    If Enlaza Then
        SQL = SQL & " codusu = " & Data1.Recordset!codusu
        SQL = SQL & " AND feccaja= '" & Format(Data1.Recordset!feccaja, FormatoFecha) & "'"
        Else
        SQL = SQL & "  codusu = -1"
    End If
    SQL = SQL & " ORDER BY slicaja.numlinea"
    MontaSQLCarga = SQL
    
    
    
    'SELECT
    'select codusu, feccaja, numlinea, slicaja.tipomovi, siglas,numserie, numfaccl, fecfaccl, numfacpr,
    'fecfacpr , slicaja.codmacta, Nommacta, slicaja.codforpa, ampliaci, ImporteD, ImporteH
    'From slicaja, stipcaja, cuentas
    'Where slicaja.tipomovi = stipcaja.tipomovi And slicaja.codmacta = cuentas.codmacta

    
End Function


Private Sub AnyadirLinea(Limpiar As Boolean)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    Linliapu = ObtenerSigueinteNumeroLinea
    'Situamos el grid al final
    DeseleccionaGrid DataGrid1
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If
    cmdAceptar.Caption = "Aceptar"
    LLamaLineas anc, 1, Limpiar
  
    'Ponemos el foco
    'en el cmbo
    cmbTipo.Enabled = True
    txtaux(0).Enabled = True
    cmdTrer(0).Visible = True
    cmdTrer(1).Visible = True
    cmbTipo.ListIndex = 0
    PonerFrameCombo 0
    HabilitarCamposFacturas True
    MaxImporteFactura = 0
    cmbTipo.SetFocus
    
End Sub

Private Sub ModificarLinea()
Dim cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    

    Linliapu = Adodc1.Recordset!NumLinea
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    DeseleccionaGrid DataGrid1
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtaux(0).Text = Adodc1.Recordset.Fields!codmacta
    txtaux(1).Text = Adodc1.Recordset.Fields!Nommacta
    txtaux(2).Text = DataGrid1.Columns(5).Text
    txtaux(3).Text = DataGrid1.Columns(6).Text

    cad = DBLet(Adodc1.Recordset.Fields!ImporteD)
    If cad <> "" Then
        txtaux(4).Text = Format(cad, "0.00")
    Else
        txtaux(4).Text = cad
    End If
    cad = DBLet(Adodc1.Recordset.Fields!ImporteH)
    If cad <> "" Then
        txtaux(5).Text = Format(cad, "0.00")
    Else
        txtaux(5).Text = cad
    End If
 
 


 
    'El combo
    cmbTipo.ListIndex = Adodc1.Recordset!tipomovi
    PonerFrameCombo CInt(Adodc1.Recordset!tipomovi)
 
 
    ''numserie, numfaccl, fecfaccl, numfacpr,fecfacpr,numvenci  field 11
    If cmbTipo.ListIndex <> 0 Then
        For I = 6 To 9
            txtaux(I).Text = ""
        Next
    Else
        txtaux(6).Text = DBLet(Adodc1.Recordset!NUmSerie, "T")
        txtaux(7).Text = DBLet(Adodc1.Recordset!numfaccl, "T")
        txtaux(8).Text = DBLet(Adodc1.Recordset!fecfaccl, "T")
        txtaux(9).Text = DBLet(Adodc1.Recordset!numvenci, "T")
    End If
 
    If cmbTipo.ListIndex <> 1 Then
        For I = 10 To 12
            txtaux(I).Text = ""
        Next
    Else
        txtaux(10).Text = DBLet(Adodc1.Recordset!numfacpr, "T")
        txtaux(11).Text = DBLet(Adodc1.Recordset!fecfacpr, "T")
        txtaux(12).Text = DBLet(Adodc1.Recordset!numvenci, "T")
    End If
 
    Text3(2).Text = DBLet(Adodc1.Recordset!nomforpa)
 
 
    'Obtengo el importe ya cobrado
    If Val(Adodc1.Recordset!tipomovi) < 2 Then
        Set miRsAux = New ADODB.Recordset
        ObtenerDatosCobroPago
        Set miRsAux = Nothing
    End If
     
    cmdTrer(0).Visible = False
    cmdTrer(1).Visible = False
    LLamaLineas anc, 2, False
    txtaux(0).Enabled = Val(Adodc1.Recordset!tipomovi) <> 1
    If Val(Adodc1.Recordset!tipomovi) < 2 Then
        Ponerfoco txtaux(2)
    Else
        Ponerfoco txtaux(0)
    End If
    cmbTipo.Enabled = False
End Sub

Private Sub EliminarLineaFactura()
Dim P As Integer

    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de caja." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar la linea: "
    SQL = SQL & DataGrid1.Columns(2).Text & " - " & DataGrid1.Columns(4).Text & " - " & DataGrid1.Columns(6).Text & " : " & Text3(5).Text
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
     
        P = Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from slicaja"
        SQL = SQL & " WHERE numlinea = " & Adodc1.Recordset!NumLinea
        SQL = SQL & " AND feccaja='" & Format(Data1.Recordset!feccaja, FormatoFecha)
        SQL = SQL & "' AND codusu=" & Data1.Recordset!codusu & ";"
        

        DataGrid1.Enabled = False
        Conn.Execute SQL
        
        
        'Si es factura miramos el ultcobro pago
        If Val(Adodc1.Recordset!tipomovi) < 2 Then
            Set miRsAux = New ADODB.Recordset
            'Obtengo el valor de lo que ponia en ultimo pago
            ObtenerDatosCobroPago
            If MaxImporteFactura <> 0 Or ImporteYaCobrado <> 0 Then EliminarModificarEnlazeCobroPago True
            Set miRsAux = Nothing
        End If
        
        
        CargaGrid (Not Data1.Recordset.EOF)
        DataGrid1.Enabled = True
        PosicionaLineas P
    End If
End Sub

Private Sub PosicionaLineas(Pos As Integer)
    On Error GoTo EPosicionaLineas
    If Pos > 1 Then
        If Pos > Adodc1.Recordset.RecordCount Then Pos = Adodc1.Recordset.RecordCount - 1
        Adodc1.Recordset.Move Pos
    End If
    
    Exit Sub
EPosicionaLineas:
    Err.Clear
End Sub

Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim Rs As ADODB.Recordset
    Dim I As Long
    
    Set Rs = New ADODB.Recordset
    SQL = "SELECT Max(numlinea) FROM slicaja"
    SQL = SQL & " WHERE codusu=" & Data1.Recordset!codusu
    SQL = SQL & " AND feccaja='" & Format(Data1.Recordset!feccaja, FormatoFecha)
    SQL = SQL & "' ;"
    Rs.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then I = Rs.Fields(0)
    End If
    Rs.Close
    ObtenerSigueinteNumeroLinea = I + 1
End Function



'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------


Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    'DeseleccionaGrid DataGrid1
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
    framelineas2.Visible = Not B
    frameextras2.Visible = B
    'Habilitamos los botones de cuenta
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B

    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim I As Integer
    
    
    DataGrid1.Enabled = Not Visible
    
    For I = 0 To txtaux.Count - 1
        txtaux(I).Visible = Visible
        If I < 6 Then txtaux(I).Top = Altura  'son los campos del grid
    Next I
    cmdAux(0).Visible = Visible
    cmdAux(0).Top = Altura
    Me.Image1(0).Enabled = Visible
    Me.Image1(1).Enabled = Visible
    Me.Image1(2).Enabled = Visible
    
    cmbTipo.Visible = Visible
    cmbTipo.Top = Altura
    If Limpiar Then
        For I = 0 To txtaux.Count - 1
            txtaux(I).Text = ""
        Next I
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtaux(Index)
   
    If Index <> 5 Then
      
        .SelStart = 0
        .SelLength = Len(.Text)
    Else
        .SelStart = Len(.Text)
    End If
End With

End Sub

Private Sub txtaux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'Esto sera k hemos pulsado el ENTER
        txtAux_LostFocus Index
        cmdAceptar_Click
    Else
 
    End If
            
'
'            'Ha pulsado F5. Ponemos linea anterior
'            Select Case KeyCode
'            Case 105
'                if inde
'            Case 116
'                'PonerLineaAnterior (Index)
'
'            Case 117
'                'F6
'                'Si es el primer campo , y ha pulsado f6
'                'cogera la linea de arriba y la pondra en los txtaux
'                txtaux(Index).Text = ""
'
'            Case Else
'                If (Shift And vbCtrlMask) > 0 Then
'                    If UCase(Chr(KeyCode)) = "B" Then
'                        'OK. Ha pulsado Control + B
'                        '----------------------------------------------------
'                        '----------------------------------------------------
'                        '
'                        ' Dependiendo de index lanzaremos una opcion uotra
'                        '
'                        '----------------------------------------------------
'
'                        'De momento solo para el 5. Cliente
'                        Select Case Index
'                        Case 4
'                            txtaux(4).Text = ""
'                            Image1_Click 1
'                        Case 8
'                            txtaux(8).Text = ""
'                            Image1_Click 2
'                        End Select
'                     End If
'                End If
'            End Select
'        End If
'    End If
End Sub


''Desplegaremos su formulario asociado
'Private Function PulsadoMas(Index As Integer, KeyAscii) As Boolean
'    Select Case Index
'    Case 0
'        'Voy a poner la modificacion del "+"
'        'Es que quiere que le mostremos su formulario de regresar
'        txtaux(0).Text = ""
'        cmdAux_Click 0
'        PulsadoMas = True
'    Case 3
'        txtaux(0).Text = ""
'        Image1_Click 0
'        PulsadoMas = True
'    Case 4
'        txtaux(4).Text = ""
'        Image1_Click 1
'        PulsadoMas = True
'
'    End Select
'
'End Function



Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        Select Case Index
        Case 0
            'Si pulsando + sobre el txtaux(0) disparamos la busqueda
            If KeyAscii = 43 Then
                KeyAscii = 0
                CadenaAmpliacion = txtaux(0).Text
                txtaux(0).Text = ""
                cmdAux_Click 0
            End If
        Case 2
            If KeyAscii = 43 Then
                KeyAscii = 0
                txtaux(2).Text = ""
                Text3(2).Text = ""
                Image1_Click 2
            End If
            
        
        End Select
        'If KeyAscii = 43 Then
        '    If PulsadoMas(Index, KeyAscii) Then KeyAscii = 0
        'End If
    End If
End Sub




Private Sub txtaux_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                'LanzaPantalla Index
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
        
    
        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
            Select Case Index
            Case 0
                
                txtaux(1).Text = ""
            Case 2
                Text3(2).Text = ""
            
            End Select
            Exit Sub
        End If
 
        
        Select Case Index
        Case 0
            RC = txtaux(0).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(0).Text = RC
                txtaux(1).Text = SQL
                RC = ""
            Else
                If InStr(1, SQL, "No existe la cuenta :") > 0 Then
    
 
                Else
                    MsgBox SQL, vbExclamation
                End If
                    
                If SQL <> "" Then
                  txtaux(0).Text = ""
                  txtaux(1).Text = ""
                  RC = "NO"
                End If
            End If
            
            If RC <> "" Then Ponerfoco txtaux(0)
            
        Case 2
        
            'FORPA
            RC = ""
            If Not IsNumeric(txtaux(Index).Text) Then
                MsgBox "Campo debe ser numérico", vbExclamation
            Else
                RC = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", txtaux(Index).Text, "N")
                If RC = "" Then
                    MsgBox "Forma pago no existe", vbExclamation
                End If
            End If
            Text3(2).Text = RC
            If RC = "" Then
                txtaux(Index).Text = ""
                Ponerfoco txtaux(Index)
            End If
        Case 4, 5
                'LOS IMPORTES
                If Not EsNumerico(txtaux(Index).Text) Then
                    MsgBox "Importes deben ser numéricos.", vbExclamation
                    txtaux(Index).Text = ""
                    Ponerfoco txtaux(Index)
                    Exit Sub
                End If
                
                
                'Es numerico
                SQL = TransformaPuntosComas(txtaux(Index).Text)
                If CadenaCurrency(SQL, Importe) Then
                    txtaux(Index).Text = Format(Importe, "0.00")
                    'Ponemos el otro campo a ""
                    If Index = 4 Then
                        txtaux(5).Text = ""
                    Else
                        txtaux(4).Text = ""
                    End If
                End If
        Case 8, 11
        
                If Not EsFechaOK(txtaux(Index)) Then
                    MsgBox "Fecha incorrecta", vbExclamation
                    txtaux(Index).Text = ""
                    Ponerfoco txtaux(Index)
                End If
  
        End Select
End Sub


Private Function AuxOK() As String
Dim Importe As Currency

    'Cuenta
    If txtaux(0).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    
    If Not IsNumeric(txtaux(0).Text) Then
        AuxOK = "Cuenta debe ser numrica"
        Exit Function
    End If
    

    
    If Not EsCuentaUltimoNivel(txtaux(0).Text) Then
        AuxOK = "La cuenta no es de último nivel"
        Exit Function
    End If
    
    

    'FORPA
    If txtaux(2).Text = "" Then
        AuxOK = "Forma de pago no puede estar vacia"
        Exit Function
    End If
        
    If txtaux(2).Text <> "" Then
        If Not IsNumeric(txtaux(2).Text) Then
            AuxOK = "La forma pago debe de ser numérica."
            Exit Function
        End If
    End If
    
    'Importe
    If txtaux(4).Text <> "" Then
        If Not EsNumerico(txtaux(4).Text) Then
            AuxOK = "El importe COBROS debe ser numérico"
            Exit Function
        End If
    End If
    
    If txtaux(5).Text <> "" Then
        If Not EsNumerico(txtaux(5).Text) Then
            AuxOK = "El importe PAGOS debe ser numérico"
            Exit Function
        End If
    End If
    
    
    If Me.cmbTipo.ListIndex < 2 Then
        I = 4
        If Me.cmbTipo.ListIndex = 1 Then I = 5
        'FRACLI FRAPRO
        If MaxImporteFactura > 0 Then
            'SIGNIFICA que la han traido
            Importe = ImporteFormateado(txtaux(I).Text)
                    
            'MaxImporteFactura
            If Importe > MaxImporteFactura Then
                AuxOK = "Importe excede el total pendiente de esta factura"
                Exit Function
            End If
        End If
    End If
    
    If Not (txtaux(4).Text = "" Xor txtaux(5).Text = "") Then
        AuxOK = "Solo el cobro, o solo el pago, tiene que tener valor"
        Exit Function
    End If
    

    
    'si es factura proveedor o factura cliente deberiamos avisar si los campos de la factura estan vacios
    SQL = ""
    If Me.cmbTipo.ListIndex = 0 Then
        'Clientes
        't 6,7,8,9
        For I = 6 To 9
            If txtaux(I).Text = "" Then SQL = "M"
        Next
    ElseIf Me.cmbTipo.ListIndex = 1 Then
        For I = 10 To 12
            If txtaux(I).Text = "" Then SQL = "M"
        Next
    End If
    If SQL <> "" Then
        
        If Me.cmbTipo.ListIndex = 0 Then
            I = 6
            SQL = "cobro"
        Else
            I = 10
            SQL = "pago"
        End If
        SQL = "Es el " & SQL & " de una factura. Deberia rellenar todos los datos de la misma. " & vbCrLf & "¿Continuar de igual modo?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
            AuxOK = "NO"
            Ponerfoco txtaux(I)
            Exit Function
        End If
    End If
    
    
    If CuentaBloqeada(txtaux(0).Text, CDate(Text1(1).Text), False) Then
        AuxOK = "Cuenta bloqueada"
        Exit Function
    End If
    
        
    
    
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS
        'INSERT INTO slicaja (codusu, feccaja, numlinea, tipomovi, numserie, numfaccl, fecfaccl, numfacpr, fecfacpr, codmacta, codforpa, numvenci, ampliaci, imported, importeh) VALUES
        SQL = "INSERT INTO slicaja (codusu, feccaja, numlinea, tipomovi, numserie, numfaccl, fecfaccl, "
        SQL = SQL & "numfacpr, fecfacpr, numvenci,codmacta, codforpa,  ampliaci, imported, importeh) VALUES ("
        
        SQL = SQL & Text1(4).Text & ",'" & Format(Data1.Recordset!feccaja, FormatoFecha) & "',"
        SQL = SQL & Linliapu & "," & cmbTipo.ItemData(Me.cmbTipo.ListIndex) & ","
        
        'Ahora SI es pago vario o traspaso los campos van a NULL
        If Me.cmbTipo.ListIndex > 1 Then
            SQL = SQL & "NULL,NULL,NULL,NULL,NULL,NULL,"
            
        Else
            'Si es factura CLIENTE o PROVEEDOR
            If cmbTipo.ListIndex = 0 Then
                'CLIENTE  txt 6,7,8,9
                
                SQL = SQL & DBSet1(txtaux(6), "T") & "," & DBSet1(txtaux(7), "N") & ","
                SQL = SQL & DBSet1(txtaux(8), "F") & ","
                'Los campos 10,11,12 a NULL
                SQL = SQL & "NULL,NULL,"
                'El nmvenci
                SQL = SQL & DBSet1(txtaux(9), "N") & ","
            Else
                'PROVEEDOR txt 10,11,12
                SQL = SQL & "NULL,NULL,NULL,"   'los 3 primeros a NULL
                SQL = SQL & DBSet1(txtaux(10), "T") & "," & DBSet1(txtaux(11), "F") & ","
                'El nmvenci
                SQL = SQL & DBSet1(txtaux(12), "N") & ","
            End If
        End If
        
        'Codmacta, codforpa,ampliaci
        SQL = SQL & "'" & txtaux(0).Text & "'," & txtaux(2).Text & ","
        SQL = SQL & DBSet1(txtaux(3), "T") & ","
        
        If txtaux(4).Text = "" Then
          SQL = SQL & "NULL," & TransformaComasPuntos(txtaux(5).Text)
          Else
          SQL = SQL & TransformaComasPuntos(txtaux(4).Text) & ",NULL"
        End If
        
        'Marca de entrada manual de datos
        SQL = SQL & ")"
        
    Else
    
        'MODIFICAR
        'slicaja (codusu, feccaja, numlinea, tipomovi, numserie, numfaccl, fecfaccl, numfacpr,
        'fecfacpr, codmacta, codforpa, numvenci, ampliaci, imported, importeh) VALUES
        SQL = "UPDATE slicaja SET "
        
        
        
        'Si es factura CLIENTE o PROVEEDOR
        If cmbTipo.ListIndex = 0 Then
            'CLIENTE  txt 6,7,8,9
            SQL = SQL & "numserie = " & DBSet1(txtaux(6), "T") & ", numfaccl = " & DBSet1(txtaux(7), "N") & ","
            SQL = SQL & "fecfaccl = " & DBSet1(txtaux(8), "F") & ","
            'El nmvenci
            SQL = SQL & "numvenci = " & DBSet1(txtaux(9), "N") & ","
        ElseIf cmbTipo.ListIndex = 1 Then
            'PROVEEDOR txt 10,11,12
            SQL = SQL & " numfacpr = " & DBSet1(txtaux(10), "T") & ",fecfacpr = " & DBSet1(txtaux(11), "T") & ","
            'El nmvenci
            SQL = SQL & "numvenci = " & DBSet1(txtaux(12), "N") & ","
        End If
        
        SQL = SQL & " codmacta = '" & txtaux(0).Text & "', codforpa = " & txtaux(2).Text & ", ampliaci="
        SQL = SQL & DBSet1(txtaux(3), "T") & ","
        

        If txtaux(4).Text = "" Then
          SQL = SQL & " imported = " & ValorNulo & "," & " importeH = " & TransformaComasPuntos(txtaux(5).Text)
          Else
          SQL = SQL & " imported = " & TransformaComasPuntos(txtaux(4).Text) & "," & " importeH = " & ValorNulo
        End If
        
        
        SQL = SQL & " WHERE numlinea = " & Linliapu
        SQL = SQL & " AND codusu=" & Data1.Recordset!codusu
        SQL = SQL & " AND feccaja='" & Format(Data1.Recordset!feccaja, FormatoFecha) & "';"
        
        
        
        
        
    End If
    Conn.Execute SQL
       
    
    If ModificandoLineas = 1 Then
        'Actualizamos la tabla scobro o spago para decir que la factura esta en caja
        If Me.cmbTipo.ListIndex < 2 Then ActualizarCobrosPagos
        
    Else
        'Modificando
        If Val(Adodc1.Recordset!tipomovi) < 2 Then
            Set miRsAux = New ADODB.Recordset
            
            'Obtengo el valor de lo que ponia en ultimo pago
            If MaxImporteFactura <> 0 Then EliminarModificarEnlazeCobroPago False
            Set miRsAux = Nothing
        End If
        
    End If
    
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, "InsertarModificar linea caja.", Err.Description
End Function
 



'Private Sub DeseleccionaGrid()
'    On Error GoTo EDeseleccionaGrid
'
'    While DataGrid1.SelBookmarks.Count > 0
'        DataGrid1.SelBookmarks.Remove 0
'    Wend
'    Exit Sub
'EDeseleccionaGrid:
'        Err.Clear
'End Sub




Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
    B = DataGrid1.Enabled
    
    DataGrid1.Enabled = False
    DoEvents
    CargaGrid2 Enlaza
    DoEvents
    DataGrid1.Enabled = B
    
End Sub

Private Function Eliminar() As Boolean
On Error GoTo FinEliminar
        Conn.BeginTrans
        SQL = " WHERE  codusu=" & Data1.Recordset!codusu
        SQL = SQL & " AND feccaja='" & Format(Data1.Recordset!feccaja, FormatoFecha) & "';"
        
        'Lineas
        Conn.Execute "Delete  from slicaja " & SQL
        
        'Cabeceras
        Conn.Execute "Delete  from scacaja " & SQL
        
                
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        
        
        
        
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function Modificar() As Boolean
Dim B1 As Boolean


    On Error GoTo EModificar
     Modificar = False
     
        '-----------------------------------------------
        ' ABRIL 2006
        '
        ' Si cambia de ejercicio le ofertaremos un nuevo numero de ASIENTO
        '
        B1 = False
        If Data1.Recordset!feccaja <> CDate(Text1(1).Text) Then
            'HAN CAMBIADO DE FECHA
            B1 = True
        End If

                    
                    
                    
   
        'Comun
        
        SQL = " WHERE  codusu =" & Data1.Recordset!codusu
        SQL = SQL & " AND feccaja='" & Format(Data1.Recordset!feccaja, FormatoFecha) & "'"
        
        If B1 Then
            'Las lineas de apuntes
            SQL = "  feccaja = '" & Format(Text1(1).Text, FormatoFecha) & "' " & SQL
            Conn.Execute "UPDATE slicaja SET " & SQL
      
        End If
        Conn.Execute "UPDATE scacaja SET " & SQL
        
  
EModificar:
        If Err.Number <> 0 Then
            MuestraError Err.Number

            Modificar = False
            B1 = False
        Else
 
            Modificar = True
        End If
        

End Function





Private Sub Ponerfoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Function RecodsetVacio() As Boolean
    RecodsetVacio = True
    If Not Adodc1.Recordset Is Nothing Then
        If Not Adodc1.Recordset.EOF Then RecodsetVacio = False
    End If
End Function



'Private Sub LanzaPantalla(Index As Integer)
'Dim miI As Integer
'        '----------------------------------------------------
'        '----------------------------------------------------
'        '
'        ' Dependiendo de index lanzaremos una opcion uotra
'        '
'        '----------------------------------------------------
'
'        'De momento solo para el 5. Cliente
'        miI = -1
'        Select Case Index
'        Case 0
'            txtaux(0).Text = ""
'            miI = 3
'        Case 3
'            txtaux(3).Text = ""
'            miI = 0
'        Case 4
'            txtaux(4).Text = ""
'            miI = 1
'
'        Case 8
'            txtaux(8).Text = ""
'            miI = 2
'        End Select
'        If miI >= 0 Then Image1_Click miI
'End Sub



Private Sub CargaCombo()
    SQL = "Select * from stipcaja  "
    If Not UsuarioCajaPredeterminada Then SQL = SQL & " WHERE tipomovi<>3"  'El usuario predeter es el unico que hace trasapasos
    SQL = SQL & " order by tipomovi  "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Me.cmbTipo.Clear
    While Not miRsAux.EOF
        cmbTipo.AddItem miRsAux!siglas
        cmbTipo.ItemData(cmbTipo.NewIndex) = miRsAux!tipomovi
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_LostFocus()
'  WheelUnHook
'End Sub


'Pruebas j

Private Sub PonerDatosFactura()
Dim I As Integer

    On Error Resume Next
    'numserie, numfaccl, fecfaccl, numfacpr,fecfacpr,numvenpr
    SQL = ""
    For I = 11 To 16
        SQL = SQL & " " & DBLet(Adodc1.Recordset.Fields(I), "T")
    Next I
    Text3(5).Text = Trim(SQL)
    Text3(4).Text = CStr(Adodc1.Recordset!nomforpa) & " (" & CStr(Adodc1.Recordset!descformapago) & ")"
    If Err.Number <> "" Then Err.Clear
End Sub


Private Sub PonerFrameCombo(Indice As Integer)
    Me.FrameCombo(0).Visible = Indice = 0 ' cmbTipo.ListIndex = 0
    Me.FrameCombo(1).Visible = Indice = 1
    If Indice < 2 Then
        'EL TEXTO
        Label2(8).Caption = ""
        HabilitarCamposFacturas False
    Else
        If Indice = 2 Then
            Label2(8).Caption = "PAGOS VARIOS"
        Else
            Label2(8).Caption = "TRASPASO CAJA"
        End If
    End If
    
End Sub



'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'
'           CONTABILIZACION DE LA CAJA
'
'
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Private Function ContabilizarCaja() As Boolean
Dim CtaPendAplicar As String
Dim Diario As Integer
Dim LaCaja As String



    ContabilizarCaja = False
    
    
  
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(*) from scacaja where  feccaja < '" & Format(Data1.Recordset!feccaja, FormatoFecha) & "'"
    SQL = SQL & " AND codusu =" & Text1(4).Text
    I = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then I = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If I > 0 Then
        MsgBox "Existe fechas anteriores a la actual pendiente de cerrar", vbExclamation
        
    Else
        'Es el que corresponde
        'Veremos los datos asociados a la caja y la cuenta
        SQL = "Select ctacaja,diario from susucaja where codusu =" & Text1(4).Text
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            LaCaja = miRsAux.Fields(0)
            Diario = miRsAux!Diario
        Else
            MsgBox "Cuenta caja no configurada para el usuario", vbExclamation
            I = 1
        End If
        miRsAux.Close
        
        
        
        
        'Ahora obtengo la cuenta de partidas pendientes de aplicacion
        
        If I = 0 Then
            CtaPendAplicar = "Select par_pen_apli from paramtesor"
            miRsAux.Open CtaPendAplicar, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            CtaPendAplicar = ""
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux.Fields(0)) Then CtaPendAplicar = miRsAux.Fields(0)
            End If
            miRsAux.Close
            
            If CtaPendAplicar = "" Then
                MsgBox "Cuenta de partidas pendientes de aplicar no configurada.", vbExclamation
                I = 1
                
            End If
            
            
        End If
        
        
    End If
    
    If I > 0 Then
        Set miRsAux = Nothing
        Exit Function
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Inserto en tmpfaclin las lineas de caja que despues tendre que ver si su cobro/PAGO
    ' hay que eliminarlo o no
    If Not PreparaTemporalEliminarCobrosPagos Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If ContabilizarCierreCajaNuevo(Data1.Recordset!feccaja, LaCaja, CtaPendAplicar, Diario) Then
        
        
        Me.Refresh
        DoEvents
        
        '****************************************************************
        'Para cada linea de caja vemos si hay que borrar los vencimientos
        Set miRsAux = New ADODB.Recordset
        ComprobarEliminarVtosFacturas
        Set miRsAux = Nothing

        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim - 1
            End If
            PonerCampos
            DataGrid1.Enabled = True
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        End If
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault

End Function




Private Sub PonerDatosPartidasPendientesAplicacion()
    On Error GoTo EDatosPartidasPendientesAplicacion
    DatosPartidasPendientesAplicacion = "Select par_pen_apli,nommacta from paramtesor,cuentas where cuentas.codmacta=paramtesor.par_pen_apli"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open DatosPartidasPendientesAplicacion, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DatosPartidasPendientesAplicacion = "||"
    If Not miRsAux.EOF Then DatosPartidasPendientesAplicacion = DBLet(miRsAux.Fields(0), "T") & "|" & DBLet(miRsAux.Fields(1), "T") & "|"
    miRsAux.Close
EDatosPartidasPendientesAplicacion:
    If Err.Number <> 0 Then Err.Clear
    Set miRsAux = Nothing
    
End Sub


Private Sub ActualizarCobrosPagos()

    On Error GoTo EActualizarCobrosPagos
    'Como las variables las voy a resetar la puedo machacar
    'Veremos el importe total a llevar a ultimo cobro/pago
    I = 5
    If Me.cmbTipo.ListIndex = 0 Then I = 4
    MaxImporteFactura = CCur(txtaux(I).Text)
    ImporteYaCobrado = ImporteYaCobrado + MaxImporteFactura
    
    
    If Me.cmbTipo.ListIndex = 0 Then
        'COBROS
        SQL = "UPDATE scobro set estacaja=1 , fecultco = '" & Format(Text1(1).Text, FormatoFecha) & "' , impcobro = " & TransformaComasPuntos(CStr(ImporteYaCobrado))

        SQL = SQL & " WHERE numserie = '" & txtaux(6).Text
        SQL = SQL & "' AND codfaccl =" & Val(txtaux(7).Text) & " AND fecfaccl = '" & Format(txtaux(8).Text, FormatoFecha) & "' AND numorden =" & txtaux(9).Text
    Else
        'PAGOS
        SQL = "UPDATE spagop set estacaja=1 , fecultpa = '" & Format(Text1(1).Text, FormatoFecha) & "' ,imppagad = " & TransformaComasPuntos(CStr(ImporteYaCobrado))
        SQL = SQL & " WHERE  numfactu ='" & DevNombreSQL(txtaux(10).Text)
        SQL = SQL & "' AND fecfactu = '" & Format(txtaux(11).Text, FormatoFecha) & "' AND numorden =" & txtaux(12).Text
        SQL = SQL & " AND ctaprove = '" & Format(txtaux(0).Text) & "'"
    End If
    Conn.Execute SQL
    
    Exit Sub
EActualizarCobrosPagos:
    MuestraError Err.Number, "Actualizar Cobros Pagos"
End Sub


Private Sub HabilitarCamposFacturas(Si As Boolean)
Dim I As Integer
    For I = 6 To 12
        txtaux(I).Enabled = Si
    Next I
End Sub

Private Sub MontaCadenaSQLObtenerCobro()

    With Adodc1.Recordset
        If Val(Adodc1.Recordset!tipomovi) = 0 Then
            SQL = " WHERE numserie = '" & !NUmSerie
            SQL = SQL & "' AND fecfaccl = '" & Format(!fecfaccl, FormatoFecha) & "'"
            SQL = SQL & " AND codfaccl = " & !numfaccl & " AND numorden = " & !numvenci
        Else
            'PAGO
            SQL = " WHERE  numfactu ='" & DevNombreSQL(!numfacpr)
            SQL = SQL & "' AND fecfactu = '" & Format(!fecfacpr, FormatoFecha) & "' AND numorden =" & !numvenci
            SQL = SQL & " AND ctaprove = '" & !codmacta & "'"
            
        End If
    End With
End Sub


Private Sub ObtenerDatosCobroPago()

    
    MaxImporteFactura = 0
    'Qui montamos el WHERE .....
    MontaCadenaSQLObtenerCobro
    'Segun sea cobro pago
    If Val(Adodc1.Recordset!tipomovi) = 0 Then
        'FRACLI
        
        SQL = "Select impvenci,impcobro,gastos FROM scobro" & SQL
    Else
        
        SQL = "Select impefect,imppagad  FROM spagop" & SQL
    
    End If
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Val(Adodc1.Recordset!tipomovi) = 0 Then
             MaxImporteFactura = DBLet(miRsAux!Gastos, "N")
             MaxImporteFactura = MaxImporteFactura + miRsAux!impvenci
             ImporteYaCobrado = DBLet(miRsAux!impcobro, "N")
             'Le quito lo que tiene de importe y asin
             ImporteYaCobrado = ImporteYaCobrado - DBLet(Adodc1.Recordset!ImporteD, "N")
             
             'Ahora fijo el MaxImporteFactura
             MaxImporteFactura = MaxImporteFactura - ImporteYaCobrado
        Else
             MaxImporteFactura = DBLet(miRsAux!ImpEfect, "N")
             ImporteYaCobrado = DBLet(miRsAux!imppagad, "N")
             'Le quito lo que tiene de importe y asin
             ImporteYaCobrado = ImporteYaCobrado - DBLet(Adodc1.Recordset!ImporteH, "N")
             
             'Ahora fijo el MaxImporteFactura
             MaxImporteFactura = MaxImporteFactura - ImporteYaCobrado
        End If
    End If
    miRsAux.Close
End Sub


Private Sub EliminarModificarEnlazeCobroPago(Eliminar As Boolean)
Dim impo As Currency


    On Error GoTo EEliminarEnlazeCobroPago
    
    
    'MODIFICAR
    If Not Eliminar Then
        
        If Val(Adodc1.Recordset!tipomovi) = 0 Then
            impo = CCur(txtaux(4).Text)
        Else
            impo = CCur(txtaux(5).Text)
        End If
        impo = impo + ImporteYaCobrado
    Else
    
'        If Val(adodc1.Recordset!tipomovi) = 0 Then
'            impo = adodc1.Recordset!ImporteD
'        Else
'            impo = adodc1.Recordset!ImporteH
'        End If
        impo = ImporteYaCobrado
        
    End If


        
    'CADENA WHERE
    MontaCadenaSQLObtenerCobro
    
    If Val(Adodc1.Recordset!tipomovi) = 0 Then
        CadenaAmpliacion = "UPDATE scobro SET impcobro = "
    Else
        CadenaAmpliacion = "UPDATE spagop SET imppagad = "
    End If
    
    If impo = 0 Then
        CadenaAmpliacion = CadenaAmpliacion & "NULL"
    Else
        CadenaAmpliacion = CadenaAmpliacion & TransformaComasPuntos(CStr(impo))
    End If
    
    If Val(Adodc1.Recordset!tipomovi) = 0 Then
        CadenaAmpliacion = CadenaAmpliacion & ", fecultco ="
    Else
        CadenaAmpliacion = CadenaAmpliacion & ", fecultpa ="
    End If
    
    If impo = 0 Then
        CadenaAmpliacion = CadenaAmpliacion & "NULL"
    Else
        CadenaAmpliacion = CadenaAmpliacion & "'" & Format(Text1(1).Text, FormatoFecha) & "'"
    End If
    
    CadenaAmpliacion = CadenaAmpliacion & ", estacaja = " & Abs(Not Eliminar)
    
        
        
    CadenaAmpliacion = CadenaAmpliacion & SQL
        
    'Ejecutamos
    Conn.Execute CadenaAmpliacion
    
    

EEliminarEnlazeCobroPago:
    If Err.Number <> 0 Then MuestraError Err.Number
    CadenaAmpliacion = ""
End Sub


Private Sub MontaSQLContabilizar(Cobros As Boolean)
    
    SQL = "SELECT * from slicaja WHERE codusu = " & Data1.Recordset!codusu
    SQL = SQL & " AND feccaja= '" & Format(Data1.Recordset!feccaja, FormatoFecha) & "' AND tipomovi = "
    If Cobros Then
        I = 0
    Else
        I = 1
    End If
    SQL = SQL & I
    
End Sub

Private Sub MontaSQLEnlaceCobroPago(ByRef Rtt As ADODB.Recordset, Cobro As Boolean)
    
    SQL = SQL & " WHERE "
    If Cobro Then
        SQL = SQL & " fecfaccl = '" & Format(Rtt!Fecha, FormatoFecha)
        SQL = SQL & "' AND numserie = '" & Rtt!Cta & "' AND codfaccl = " & Rtt!numfac
     Else
     
        SQL = SQL & "  fecfactu = '" & Format(Rtt!Fecha, FormatoFecha)
        SQL = SQL & "' AND ctaprove = '" & Rtt!Cta & "' AND numfactu = '" & Rtt!numfac & "'"
     End If
    SQL = SQL & " AND numorden = " & Rtt!NIF
    SQL = SQL & " AND estacaja = 1"
     
End Sub

Private Sub ComprobarEliminarVtosFacturas()
Dim RT As ADODB.Recordset

    On Error GoTo EComprobarEliminarVtosFacturas
    Set RT = New ADODB.Recordset
    
    
    
    'Cobros
    Me.lblIndicador.Caption = "Vtos fac. cliente"
    Me.lblIndicador.Refresh
    SQL = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " AND IVA = '0'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = "Select impvenci, impcobro, gastos from scobro "
        MontaSQLEnlaceCobroPago miRsAux, True
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            'YA tengo el VTO. Veremos si me lo tengo que cargar o no
            ImporteYaCobrado = RT!impvenci + DBLet(RT!Gastos, "N")
            ImporteYaCobrado = ImporteYaCobrado - DBLet(RT!impcobro, "N")
            
            'Si es 0 significa que ya esta cobrad totalmente y lo eliminare
            'En caso contrario le quitare la marca de "esta en caja"
            If ImporteYaCobrado = 0 Then
                SQL = "DELETE from scobro "
    
            Else
                SQL = "UPDATE scobro set estacaja = 0 "
            End If
            MontaSQLEnlaceCobroPago miRsAux, True
            Ejecuta SQL
        End If
        RT.Close
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'Pagos
    Me.lblIndicador.Caption = "Vtos fac. proveedor"
    Me.lblIndicador.Refresh
    SQL = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " AND IVA = '1'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = "Select impefect, imppagad from spagop "
        MontaSQLEnlaceCobroPago miRsAux, False
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            'YA tengo el VTO. Veremos si me lo tengo que cargar o no
            ImporteYaCobrado = RT!ImpEfect
            ImporteYaCobrado = ImporteYaCobrado - DBLet(RT!imppagad, "N")
            
            'Si es 0 significa que ya esta cobrad totalmente y lo eliminare
            'En caso contrario le quitare la marca de "esta en caja"
            If ImporteYaCobrado = 0 Then
                SQL = "DELETE from spagop "
    
            Else
                SQL = "UPDATE spagop set estacaja = 0 "
            End If
            MontaSQLEnlaceCobroPago miRsAux, False
            Ejecuta SQL
        End If
        RT.Close

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    

    Exit Sub
EComprobarEliminarVtosFacturas:
    MuestraError Err.Description, "Comprobar Eliminar Vtos Facturas"
  
End Sub




Private Function PreparaTemporalEliminarCobrosPagos() As Boolean
Dim cad As String
    On Error GoTo EPreparaTemporalEliminarCobrosPagos
    PreparaTemporalEliminarCobrosPagos = False
    
    Set miRsAux = New ADODB.Recordset
    Conn.Execute "DELETE FROM tmpfaclin where codusu = " & vUsu.Codigo
    
    
    
    NumRegElim = 0
    MontaSQLContabilizar True
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    
    
    
    SQL = ""
    While Not miRsAux.EOF
        If Not (IsNull(miRsAux!numfaccl) Or IsNull(miRsAux!fecfaccl) Or IsNull(miRsAux!NUmSerie) Or IsNull(miRsAux!numvenci)) Then
            NumRegElim = NumRegElim + 1
            SQL = SQL & ",(" & vUsu.Codigo & "," & NumRegElim & ",0,'" & miRsAux!numfaccl & "','" & miRsAux!NUmSerie & "','"
            SQL = SQL & Format(miRsAux!fecfaccl, FormatoFecha) & "','" & miRsAux!numvenci & "',"
            'los importes
            SQL = SQL & TransformaComasPuntos(CStr(miRsAux!ImporteD)) & ",0,0)"
            
    
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
                'Quito la primera coma
    If SQL <> "" Then cad = Mid(SQL, 2)
    
    
    MontaSQLContabilizar False
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = ""
    While Not miRsAux.EOF
        If Not (IsNull(miRsAux!codmacta) Or IsNull(miRsAux!fecfacpr) Or IsNull(miRsAux!numfacpr) Or IsNull(miRsAux!numvenci)) Then
            NumRegElim = NumRegElim + 1
            SQL = SQL & ",(" & vUsu.Codigo & "," & NumRegElim & ",1,'" & DevNombreSQL(miRsAux!numfacpr) & "','" & miRsAux!codmacta & "','"
            SQL = SQL & Format(miRsAux!fecfacpr, FormatoFecha) & "','" & miRsAux!numvenci & "',"
            'los importes
            SQL = SQL & TransformaComasPuntos(CStr(miRsAux!ImporteH)) & ",0,0)"
    
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If SQL <> "" Then
        If cad = "" Then SQL = Mid(SQL, 2)
        cad = cad & SQL
    End If
    
    If cad <> "" Then
        SQL = "INSERT INTO tmpfaclin (codusu,codigo,IVA,Numfac,cta,Fecha,NIF,Total,Imponible,ImpIVA) VALUES " & cad
        Conn.Execute SQL
    End If
    SQL = ""
    Set miRsAux = Nothing
    NumRegElim = 0
    PreparaTemporalEliminarCobrosPagos = True
    Exit Function
EPreparaTemporalEliminarCobrosPagos:
    MuestraError Err.Number, "Preparando datos temporal contabilizacion caja"
End Function
