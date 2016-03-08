VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAJA"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmCaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameMante 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7440
         Picture         =   "frmCaja.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   5040
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmCaja.frx":1A7E
         Left            =   960
         List            =   "frmCaja.frx":1A8E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.Frame FrMantimiento 
         Height          =   5175
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   7935
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Text            =   "Text3"
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   4680
            TabIndex        =   8
            Text            =   "Text2"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CommandButton cmdTraspaso 
            Caption         =   "Traspasar"
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   9
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "La retirada de efectivo se marca con importe en negativo."
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   51
            Top             =   2880
            Width           =   4095
         End
         Begin VB.Label Label4 
            Caption         =   "Ampliación"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   50
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label Label4 
            Caption         =   "Importe "
            Height          =   195
            Index           =   5
            Left            =   4680
            TabIndex        =   49
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label Label6 
            Caption         =   "Traspaso interno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   6015
         End
      End
      Begin VB.Frame FrMantimiento 
         Height          =   5175
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   7935
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   4800
            TabIndex        =   11
            Text            =   "Text2"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   10
            Text            =   "Text3"
            Top             =   1680
            Width           =   3975
         End
         Begin VB.CommandButton cmdTraspaso 
            Caption         =   "Pagar"
            Height          =   375
            Index           =   0
            Left            =   5520
            TabIndex        =   12
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "SUMINISTRADOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   55
            Top             =   600
            Width           =   6015
         End
         Begin VB.Label Label4 
            Caption         =   "Importe "
            Height          =   195
            Index           =   8
            Left            =   4800
            TabIndex        =   54
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label Label4 
            Caption         =   "Ampliación"
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   53
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label Label9 
            Caption         =   "La retirada de efectivo se marca con importe en negativo."
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   52
            Top             =   2880
            Width           =   4095
         End
      End
      Begin VB.Frame FrMantimiento 
         Height          =   5175
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   7935
         Begin VB.TextBox txtCta 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox DtxtCta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   21
            Text            =   "Text5"
            Top             =   360
            Width           =   2715
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4335
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   7646
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Numero"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fec Fac."
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Orden"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Fec. Vto"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Cobrado"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Importe"
               Object.Width           =   2469
            EndProperty
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   1
            Left            =   1080
            Picture         =   "frmCaja.frx":1AD3
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame FrMantimiento 
         Height          =   5175
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   7935
         Begin VB.TextBox DtxtCta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   22
            Text            =   "Text5"
            Top             =   240
            Width           =   2715
         End
         Begin VB.TextBox txtCta 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4335
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   7646
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Serie"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Numero"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Fec Fac."
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Orden"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Fec. Vto"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Cobrado"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Importe"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   585
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   0
            Left            =   840
            Picture         =   "frmCaja.frx":1BD5
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Leyendo ..."
         Height          =   255
         Left            =   6360
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4800
         Picture         =   "frmCaja.frx":1CD7
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblFecha 
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
         Index           =   0
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrameEfectua 
      Height          =   5655
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelarPago 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6600
         TabIndex        =   35
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdEfectuarPago 
         Caption         =   "Efectuar"
         Height          =   375
         Left            =   5160
         TabIndex        =   34
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2400
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Label Label4 
         Caption         =   "Importe "
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   39
         Top             =   3600
         Width           =   1380
      End
      Begin VB.Label Label5 
         Caption         =   "FECHA "
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Importe abonado"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   33
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Importe VTO"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   32
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "FACTURA"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   2160
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente / Proveedor"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   7410
      End
   End
   Begin VB.Frame FrameVer 
      Height          =   7575
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   9495
      Begin VB.OptionButton optVer 
         Caption         =   "Suministros y traspasos"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optVer 
         Caption         =   "Pagos proveedores"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   44
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optVer 
         Caption         =   "Cobros clientes"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   360
         Width           =   6375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6015
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuenta"
            Object.Width           =   2222
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Efecto"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   6600
         Picture         =   "frmCaja.frx":1DD9
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Leyendo datos desde BD"
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
         Left            =   6840
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Frame FrameCierreCaja 
      Height          =   2895
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   5655
      Begin VB.CheckBox ChkCierre 
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdCierre 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   61
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdCierre 
         Caption         =   "Efectuar cierre"
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   60
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hacer cierre de caja"
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
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   4650
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Hasta fecha"
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
         Left            =   840
         TabIndex        =   58
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   2160
         Picture         =   "frmCaja.frx":2363
         Top             =   1320
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
Public LaCuentaDeCaja As String

Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1



Dim CtaDevuelta As String
Dim PrimeraVez As Boolean
Dim SQL As String
Dim ItmX As ListItem
Dim Importe As Currency


Private Sub cmdCancelarPago_Click()
    BloqueoManual False, "CajaPA_CO" & Combo1.ListIndex, ""
    PagosFrames False
End Sub

Private Sub cmdCierre_Click(Index As Integer)
    If Index = 0 Then
        If txtFecha(1).Text = "" Then Exit Sub
        
        
        
        
        
        
        
        SQL = "Seguro que desea efctuar el cierre de caja hasta la fecha " & txtFecha(1).Text & "?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        
        If HacerCierreCaja Then
            SQL = ""
        Else
            SQL = "NO"
        End If
        Screen.MousePointer = vbDefault
        If SQL <> "" Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdEfectuarPago_Click()


    'FECHA DENTRO DE EJERCICIOS
    If FechaCorrecta(CDate(txtFecha(0).Text)) > 1 Then
        MsgBox "Fecha fuera de ejercicios contables", vbExclamation
        Exit Sub
    End If
    
    
    If Text2(0).Text = "" Then
        MsgBox "Introduzca el importe a pagar", vbExclamation
        Exit Sub
    End If

    
    
    
    Importe = ImporteFormateado(Text2(0).Text)
    If Importe = 0 Then
        MsgBox "El importe debe ser distinto de 0", vbExclamation
        Exit Sub
    End If
    
    'Cuanto queda por pagar
    Importe = CCur(Text2(0).Tag) - Importe

    If Importe < 0 Then
        SQL = "El importe sobrepasa el total a pagar: " & Abs(Importe) & vbCrLf & "Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If
    
    If Combo1.ListIndex = 1 Then
    
        With ListView1(1).SelectedItem
            
            'INSERTAMOS EN CAJA
            '----------------------------------------------------------------------------------------
            SQL = "INSERT INTO scaja (feccaja, ctacaja, numserie, codmacta, numfactu, fecfactu,"
            SQL = SQL & "numorden, fecefect, impefect) VALUES ('"
            'VAlores sql
            SQL = SQL & Format(Me.txtFecha(0).Text, FormatoFecha) & "','" & LaCuentaDeCaja & "','"
            
            'LOS VALORES DE LA FACTURA. La serie para proveedores se marca con el 2
            SQL = SQL & "2','" & txtCta(1).Text & "','" & .Text & "',"
            SQL = SQL & "'" & Format(.SubItems(1), FormatoFecha) & "',"
            SQL = SQL & .SubItems(2) & ",'" & Format(.SubItems(3), FormatoFecha) & "',"
            
            'Emporte
            Importe = ImporteFormateado(Text2(0).Text)
            SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ")"
            If EjecutarSQL(SQL) Then
        
                'UPDATEAMOS EL COBRO. Poniendo contdocu a 1, para no volver a cargarlo
                Importe = Importe + ImporteFormateado(Text1(3).Text)
                SQL = "UPDATE spagop set contdocu= "
                
                If ImporteFormateado(Text1(2).Text) > Importe Then
                     'Aun no esta pagado del todo
                     SQL = SQL & "0"
                Else
                     'YA ESTA TOTALMENTE PAGADO
                     SQL = SQL & "1"
                End If
            
                SQL = SQL & " , imppagad= " & TransformaComasPuntos(CStr(Importe))
                SQL = SQL & " ,fecultpa = '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
                SQL = SQL & " WHERE numfactu = " & Val(.Text) & " AND ctaprove = '" & txtCta(1).Text
                SQL = SQL & "' AND fecfactu = '" & Format(.SubItems(1), FormatoFecha) & "' AND numorden =" & .SubItems(2)
                EjecutarSQL SQL
            End If
            'Borramos el ITEM
            ListView1(1).ListItems.Remove ListView1(1).SelectedItem.Index
            Set ListView1(1).SelectedItem = Nothing
        End With
    Else
            'COBRO
        
            With ListView1(0).SelectedItem
               
               'INSERTAMOS EN CAJA
               '----------------------------------------------------------------------------------------
               SQL = "INSERT INTO scaja (feccaja, ctacaja, numserie, codmacta, numfactu, fecfactu,"
               SQL = SQL & "numorden, fecefect, impefect) VALUES ('"
               'VAlores sql
               SQL = SQL & Format(Me.txtFecha(0).Text, FormatoFecha) & "','" & LaCuentaDeCaja & "',"
               
               'LOS VALORES DE LA FACTURA
               SQL = SQL & "'" & .Text & "','" & txtCta(0).Text & "','" & .SubItems(1) & "',"
               SQL = SQL & "'" & Format(.SubItems(2), FormatoFecha) & "',"
               SQL = SQL & .SubItems(3) & ",'" & Format(.SubItems(4), FormatoFecha) & "',"
               
               'Emporte
               Importe = ImporteFormateado(Text2(0).Text)
               SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ")"
               If EjecutarSQL(SQL) Then
                
           
                    'UPDATEAMOS EL COBRO. Poniendo contdocu a 1, para no volver a cargarlo
                    Importe = Importe + ImporteFormateado(Text1(3).Text)
                    SQL = "UPDATE scobro set contdocu= "
                    'Si el cobro es por el total marco que esta cobrado
                    If ImporteFormateado(Text1(2).Text) > Importe Then
                         'Aun no esta pagado del todo
                         SQL = SQL & "0"
                    Else
                         'YA ESTA TOTALMENTE PAGADO
                         SQL = SQL & "1"
                    End If
                    SQL = SQL & " ,impcobro= " & TransformaComasPuntos(CStr(Importe))
                    SQL = SQL & " ,fecultco = '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
                    SQL = SQL & " WHERE numserie = '" & .Text & "' AND codfaccl = " & Val(.SubItems(1))
                    SQL = SQL & " AND fecfaccl = '" & Format(.SubItems(2), FormatoFecha) & "' AND numorden =" & .SubItems(3)
                    EjecutarSQL SQL
             End If
            'Borramos el ITEM
            ListView1(0).ListItems.Remove ListView1(0).SelectedItem.Index
            Set ListView1(0).SelectedItem = Nothing
        End With
    End If
    PagosFrames False
    BloqueoManual False, "CajaPA_CO" & Combo1.ListIndex, ""
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTraspaso_Click(Index As Integer)

    If Text2(1 + Index).Text = "" Then
        MsgBox "Introduzca el importe", vbExclamation
        Exit Sub
    End If
    Importe = ImporteFormateado(Text2(1 + Index).Text)
    If Importe = 0 Then
        MsgBox "Introduzca el importe", vbExclamation
        Exit Sub
    End If
    
    'Compruebo la caja predeterminada
    If Index = 1 Then
        SQL = "Select * from usucaja where predeterminado=1"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If Not miRsAux.EOF Then SQL = DBLet((miRsAux!CtaCaja), "T")
        miRsAux.Close
        Set miRsAux = Nothing
        If SQL = "" Then
            MsgBox "Ninguna caja marcada a tal efecto", vbExclamation
            Exit Sub
        End If
        
        If SQL = LaCuentaDeCaja Then
            MsgBox "La cuenta de caja es la misma que la marcada para traspasos", vbExclamation
            Exit Sub
        End If
        Text3(Index).Tag = SQL
    Else
        Text3(Index).Tag = LaCuentaDeCaja
    End If
    
    Importe = ImporteFormateado(Text2(1 + Index).Text)
    
    'BLOQQUEO, por si las moscas
    If BloqueoManual(True, "CajaTRASPASO", LaCuentaDeCaja) Then
    
        SQL = "Select max(numorden) from scaja where "
        SQL = SQL & " feccaja ='" & Format(txtFecha(0).Text, FormatoFecha)
        SQL = SQL & "' AND ctacaja ='" & LaCuentaDeCaja
        SQL = SQL & "' AND numserie ='" & Index & "'"   '1.- TRASPASOS '0.-SUMINISTROS
        'VALORES FIJOS
        SQL = SQL & " AND numfactu=0  AND  fecfactu='1972-04-12' "
        
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = "1"
        If Not miRsAux.EOF Then SQL = DBLet((miRsAux.Fields(0)), "N") + 1
        miRsAux.Close
        Set miRsAux = Nothing
    
    
    
    
    
                            
             ' numorden, fecefect, impefect,ampliacion
        SQL = SQL & ",'1974-05-01'," & TransformaComasPuntos(CStr(Importe)) & ",'" & DevNombreSQL(Mid(Text3(Index).Text, 1, 30)) & "')"
            'numseri codmacta numfactu, fecfactu,
        SQL = "'" & Index & "','" & Text3(Index).Tag & "','0','1972-04-12'," & SQL
                    'FECHA CAJA
        SQL = Format(Me.txtFecha(0).Text, FormatoFecha) & "','" & LaCuentaDeCaja & "'," & SQL
        
        SQL = "numorden, fecefect, impefect,ampliacion) VALUES ('" & SQL
        SQL = "INSERT INTO scaja (feccaja, ctacaja, numserie, codmacta, numfactu, fecfactu," & SQL
        
        
        Conn.Execute SQL
        
        If Index = 0 Then
            SQL = "Pago efectuado."
        Else
            SQL = "Traspaso realizado"
        End If
        MsgBox SQL, vbExclamation
        Text2(1 + Index).Text = ""
        Text3(Index).Text = ""
        
        'Desbloquear
        BloqueoManual False, "CajaTRASPASO", ""
    Else
        MsgBox "Registro bloqueado por otro usuario", vbExclamation
    End If
End Sub

Private Sub Combo1_Click()
    If Not PrimeraVez Then
        Screen.MousePointer = vbHourglass
        If Combo1.ListIndex <> Combo1.Tag Then
            ActivarFrames
            Combo1.Tag = Combo1.ListIndex
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Combo1_GotFocus()
    Combo1.Tag = Combo1.ListIndex
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Screen.MousePointer = vbHourglass
    If Combo1.ListIndex <> Combo1.Tag Then ActivarFrames
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            Combo1.SetFocus
        Case 1
            CargaManteCobrosPagos
        End Select
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    FrameEfectua.Visible = False
    FrameVer.Visible = False
    frameMante.Visible = False
    FrameCierreCaja.Visible = False
    Select Case Opcion
    Case 0
        Caption = "EFCTUAR COBROS / PAGOS  POR CAJA"
        Limpiar Me
        Me.txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        Combo1.ListIndex = 0
        Combo1.Tag = -1
        Caption = "CAJA.  " & vUsu.Nombre
        frameMante.Visible = True
        ActivarFrames
        CargaIconoListview Me.ListView1(0)
        CargaIconoListview Me.ListView1(1)
        Me.Width = Me.frameMante.Width + 100
        Me.Height = Me.frameMante.Height + 400
    Case 1
        'VER CAJA.
        FrameVer.Visible = True
        Caption = "MANTENIMIENTO CAJA"
        CargaComboMantenimiento
        CargaIconoListview Me.ListView1(2)
    Case 2
        FrameCierreCaja.Visible = True
        Caption = "CIERRE CAJA. " & vUsu.Nombre & " - " & LaCuentaDeCaja
        Me.Width = Me.FrameCierreCaja.Width + 100
        Me.Height = Me.FrameCierreCaja.Height + 400
        txtFecha(1).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
        cmdCierre(1).Cancel = True
    End Select
End Sub


Private Sub KEYpress(ByRef Tecla As Integer)
    If Tecla = 13 Then
        Tecla = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub ActivarFrames()
Dim I As Integer
    For I = 0 To 3
        FrMantimiento(I).Visible = (I = Combo1.ListIndex)
    Next I
   
End Sub





Private Sub frmC_Selec(vFecha As Date)
    Me.txtFecha(CInt(Me.Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    CtaDevuelta = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub Image1_Click()
Dim I  As Integer
    If ListView1(2).ListItems.Count = 0 Then Exit Sub
    If ListView1(2).SelectedItem Is Nothing Then Exit Sub
    
    

    'seguro k....
    SQL = "Va a  eliminar el pago/cobro:" & vbCrLf & vbCrLf
    SQL = SQL & "Fecha caja: " & ListView1(2).SelectedItem.Text & vbCrLf
    
    
    For I = 1 To ListView1(2).ColumnHeaders.Count - 1
        SQL = SQL & ListView1(2).ColumnHeaders(I + 1).Text & ": " & ListView1(2).SelectedItem.SubItems(I) & vbCrLf
    Next I
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    'VALE. A eliminar
    
    'ItmX.Tag = !codmacta & "|" & !NUmSerie & "|" & !numfactu & "|" & !numorden & "|" & !fecfactu & "|"
    With ListView1(2).SelectedItem
        
        SQL = Combo2.List(Combo2.ListIndex)
        SQL = Trim(Mid(SQL, 1, InStr(1, SQL, " -")))
        SQL = "' AND ctacaja = '" & SQL & "'"

        SQL = "DELETE from scaja WHERE feccaja='" & Format(.Text, FormatoFecha) & SQL
        
        SQL = SQL & " AND codmacta='" & RecuperaValor(.Tag, 1)
        SQL = SQL & "' AND numfactu='" & RecuperaValor(.Tag, 3)
        SQL = SQL & "' AND fecfactu='" & Format(RecuperaValor(.Tag, 5), FormatoFecha)
        SQL = SQL & "' AND numorden=" & RecuperaValor(.Tag, 4)
        SQL = SQL & " AND numserie='" & RecuperaValor(.Tag, 2)
        SQL = SQL & "';"
        
    End With
    Conn.Execute SQL
    
    
    'Si la opcion es client/provee entonces desactualizmos el pago del vto
    Importe = ImporteFormateado(ListView1(2).SelectedItem.SubItems(4))
    '!codmacta & "|" & !NUmSerie & "|" & !numfactu & "|" & !numorden & "|" & !fecfactu & "|"
    SQL = ""
    If Me.optVer(0).Value Then
        'Tengo k acturlizar scobros, descontando el pago y poniendo contdocu a 0
        SQL = "UPDATE scobro SET contdocu=0, impcobro=impcobro - " & TransformaComasPuntos(CStr(Importe))
        'WHERE
        SQL = SQL & " WHERE numserie = '" & RecuperaValor(ListView1(2).SelectedItem.Tag, 2)
        SQL = SQL & "' AND codfaccl = " & Val(RecuperaValor(ListView1(2).SelectedItem.Tag, 3))
        SQL = SQL & " and fecfaccl = '" & Format(RecuperaValor(ListView1(2).SelectedItem.Tag, 5), FormatoFecha)
        SQL = SQL & "' and numorden = " & RecuperaValor(ListView1(2).SelectedItem.Tag, 4)
        Conn.Execute SQL
        
    Else
        If Me.optVer(1).Value Then
            'Tengo k acturlizar spagop, descontando el pago y poniendo contdocu a 0
            SQL = "UPDATE spagop SET contdocu=0, imppagad=imppagad - " & TransformaComasPuntos(CStr(Importe))
            SQL = SQL & " WHERE ctaprove = '" & RecuperaValor(ListView1(2).SelectedItem.Tag, 1)
            SQL = SQL & "' AND numfactu = '" & RecuperaValor(ListView1(2).SelectedItem.Tag, 3)
            SQL = SQL & "' and fecfactu = '" & Format(RecuperaValor(ListView1(2).SelectedItem.Tag, 5), FormatoFecha)
            SQL = SQL & "' and numorden = " & RecuperaValor(ListView1(2).SelectedItem.Tag, 4)
            Conn.Execute SQL
        End If
    End If
    
    
    
    ListView1(2).ListItems.Remove ListView1(2).SelectedItem.Index
End Sub

Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    CtaDevuelta = ""
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
    If CtaDevuelta <> "" Then
        Me.Refresh
        txtCta(Index).Text = CtaDevuelta
        txtCta_LostFocus Index
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Set frmC = New frmCal
    Me.Tag = Index
    frmC.Fecha = Now
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub ListView1_DblClick(Index As Integer)
    If ListView1(Index).ListItems.Count = 0 Then Exit Sub
    If ListView1(Index).SelectedItem Is Nothing Then Exit Sub
    
        
    If Index < 2 Then
        PonerDatosPagoCobro
    End If
'
'Dim i As Integer
'Dim cad As String
'    cad = ""
'    For i = 1 To ListView1(Index).ColumnHeaders.Count
'        cad = cad & ListView1(Index).ColumnHeaders(i).Text & ": " & ListView1(Index).ColumnHeaders(i).Width & vbCrLf
'    Next i
'    MsgBox cad
End Sub


Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick Index
End Sub



Private Sub optVer_Click(Index As Integer)
    CargaManteCobrosPagos
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    If Not IsNumeric(Text2(Index).Text) Then
        MsgBox "Campo numerico: " & Text2(Index).Text, vbExclamation
        Text2(Index).Text = ""
        Exit Sub
    End If
    
    If InStr(1, Text2(Index).Text, ",") > 0 Then
        'ESTA FORMATEADO
        
    Else
    
        'hay k formatear
        Text2(Index).Text = Format(TransformaPuntosComas(Text2(Index).Text), FormatoImporte)
    End If
        
    
End Sub




Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'txtCta
Private Sub txtCta_GotFocus(Index As Integer)
    With txtCta(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte

    txtCta(Index).Text = Trim(txtCta(Index).Text)
  
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
        ListView1(Index).ListItems.Clear
        Exit Sub
    End If
    
    
    
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

    'Cargamos el list
    If Index < 2 Then
        Screen.MousePointer = vbHourglass
        ListView1(Index).ListItems.Clear
        Set miRsAux = New ADODB.Recordset
        If DtxtCta(Index).Text <> "" Then
            
            CargaCobrosPendientes
        End If
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Sub CargaCobrosPendientes()
    Label8.Visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Combo1.ListIndex = 0 Then
        CargaCobros
    Else
        If Combo1.ListIndex = 1 Then CargaPagos
    End If
    Label8.Visible = False
    Screen.MousePointer = vbDefault
End Sub




Private Sub CargaCobros()

    SQL = "select numserie,codfaccl,fecfaccl,numorden,fecvenci,impvenci,impcobro from scobro,sforpa,stipoformapago "
    SQL = SQL & " WHERE scobro.codforpa = sforpa.codforpa And sforpa.tipforpa = stipoformapago.tipoformapago"
    SQL = SQL & " and scobro.codmacta='" & txtCta(0).Text & "' and stipoformapago.tipoformapago = " & vbEfectivo
    'FALTARA VER LA MARCA DE SI YA ESTA COBRADO . Contdocu=1 esta ya abonado en caja
    SQL = SQL & " and contdocu = 0"
    'EL ORDEN
    SQL = SQL & " ORDER BY 1,2,3,4"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = ListView1(0).ListItems.Add
        ItmX.Text = miRsAux!NUmSerie
        ItmX.SubItems(1) = Format(miRsAux!codfaccl, "000000000")
        ItmX.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
        ItmX.SubItems(3) = miRsAux!numorden
        ItmX.SubItems(4) = Format(miRsAux!fecvenci, "dd/mm/yyyy")
        If IsNull(miRsAux!impcobro) Then
            ItmX.SubItems(5) = "0,00"
        Else
            ItmX.SubItems(5) = Format(miRsAux!impcobro, FormatoImporte)
        End If
        ItmX.SubItems(6) = Format(miRsAux!impvenci, FormatoImporte)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub



Private Sub CargaPagos()
    
    
    SQL = "select ctaprove,numfactu,fecfactu,numorden,fecultpa,imppagad,impefect,fecefect from spagop,sforpa,stipoformapago "
    SQL = SQL & " WHERE spagop.codforpa = sforpa.codforpa And sforpa.tipforpa = stipoformapago.tipoformapago"
    SQL = SQL & " and spagop.ctaprove='" & txtCta(1).Text & "' and stipoformapago.tipoformapago = " & vbEfectivo
    'FALTARA VER LA MARCA DE SI YA ESTA COBRADO . Contdocu=1 esta ya abonado en caja
    'SQL = SQL & " and contdocu = 0"
    'EL ORDEN
    SQL = SQL & " ORDER BY 1,2,3,4"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = ListView1(1).ListItems.Add
        'ItmX.Text = Format(miRsAux!numfactu, "000000000")
        ItmX.Text = miRsAux!numfactu
        ItmX.SubItems(1) = Format(miRsAux!fecfactu, "dd/mm/yyyy")
        ItmX.SubItems(2) = miRsAux!numorden
        ItmX.SubItems(3) = Format(miRsAux!fecefect, "dd/mm/yyyy")
        If IsNull(miRsAux!imppagad) Then
            ItmX.SubItems(4) = "0,00"
        Else
            ItmX.SubItems(4) = Format(miRsAux!imppagad, FormatoImporte)
        End If
        ItmX.SubItems(5) = Format(miRsAux!impefect, FormatoImporte)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    With txtFecha(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub PagosFrames(HabilitarPago As Boolean)
    Me.frameMante.Visible = Not HabilitarPago
    Me.FrameEfectua.Visible = HabilitarPago
    If HabilitarPago Then
        Me.cmdCancelarPago.Cancel = True
    Else
        Me.cmdSalir.Cancel = True
    End If
    

End Sub

Private Sub PonerDatosPagoCobro()
    
    SQL = ""
    If Combo1.ListIndex = 0 Then
        
            'FACTURA
        SQL = ListView1(0).SelectedItem.Text & " / " & ListView1(0).SelectedItem.SubItems(1) & " - " & ListView1(0).SelectedItem.SubItems(3) & "-" & ListView1(0).SelectedItem.SubItems(2)
    
    Else
        If Combo1.ListIndex = 1 Then _
         SQL = txtCta(1).Text & "|" & ListView1(1).SelectedItem.Text & " - " & ListView1(1).SelectedItem.SubItems(1) & "-" & ListView1(1).SelectedItem.SubItems(2)
        
    End If
    If SQL <> "" Then
        If Not BloqueoManual(True, "CajaPA_CO" & Combo1.ListIndex, SQL) Then
            MsgBox "Bloqueado", vbExclamation
            Exit Sub
        End If
        
        
        
    End If
    PagosFrames True
    Text1(4).Text = Me.txtFecha(0).Text
    Select Case Combo1.ListIndex
    Case 0
        Label2(4).Caption = "Efectuar cobro cliente"
         'COBRO COBRO COBRO COBRO
        Text1(0).Text = txtCta(0).Text & "  -  " & Me.DtxtCta(0).Text
        With ListView1(0).SelectedItem
            'FACTURA
             Text1(1).Text = .Text & " / " & .SubItems(1) & " - " & .SubItems(3) & "    F. Fact: " & .SubItems(2)
             
             'IMPORTES
             Text1(2).Text = .SubItems(6)
             
             Text1(3).Text = .SubItems(5)
             
             Importe = ImporteFormateado(.SubItems(6)) - ImporteFormateado(.SubItems(5))
             Text2(0).Tag = Importe
             Text2(0).Text = Format(Importe, FormatoImporte)
             Text2(0).SetFocus
             
        End With
         
         '----- FIN COBRO
            
    Case 1
        ' PAGO PAGO PAGO
        Label2(4).Caption = "Efectuar pago proveedor"
        
        Text1(0).Text = txtCta(1).Text & "  -  " & Me.DtxtCta(1).Text
        With ListView1(1).SelectedItem
            'FACTURA
             Text1(1).Text = .Text & " - " & .SubItems(2) & "    F. Fact: " & .SubItems(1)
             
             'IMPORTES
             Text1(2).Text = .SubItems(5)
             
             Text1(3).Text = .SubItems(4)
             
             Importe = ImporteFormateado(.SubItems(5)) - ImporteFormateado(.SubItems(4))
             Text2(0).Tag = Importe
             Text2(0).Text = Format(Importe, FormatoImporte)
             Text2(0).SetFocus
             
        End With
        
        
        
        
        
        
        
        
        '-------------------- FIN PAGO
        
    Case 2
        Label2(4).Caption = "Efectuar pago suministro"
    
    Case 3
        Label2(4).Caption = "Traspaso caja"
            

    End Select
    
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub


Private Sub CargaComboMantenimiento()

    SQL = "Select  ctacaja, nommacta "
    SQL = SQL & " from usucaja,cuentas "
    SQL = SQL & " WHERE ctacaja = cuentas.codmacta"
    'Si llega hasta esta pantalla es pq:
    '       .- Tiene permiso
    '       .- Es un usuario de cja
    '           Si es un usuario tendra su LaCuentaDeCaja
    '           Si no es k es el superusuario
    If LaCuentaDeCaja = "S" Then LaCuentaDeCaja = "" 'ES EL SUPERUSUARIO CAJA
    If LaCuentaDeCaja <> "" Then SQL = SQL & " and ctacaja = '" & LaCuentaDeCaja & "'"
    
    
    Combo2.Clear
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo2.AddItem miRsAux!CtaCaja & " - " & miRsAux!Nommacta
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
End Sub



Private Sub CargaManteCobrosPagos()
    
    Label7.Visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    CargaRealManteCobrosPagos
    Label7.Visible = False
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaRealManteCobrosPagos()
Dim Tipo As Byte
    ListView1(2).ListItems.Clear
    
    If Combo2.ListIndex < 0 Then Exit Sub
            
    'La cuenta actual
    SQL = Combo2.List(Combo2.ListIndex)
    SQL = Trim(Mid(SQL, 1, InStr(1, SQL, " -")))
    SQL = " AND ctacaja = '" & SQL & "'"
            
    SQL = "Select scaja.*,nommacta from scaja,cuentas where scaja.codmacta=cuentas.codmacta" & SQL
    
    
    'Ahora vamos con las cobros , pagos etc
    If optVer(0).Value Then
        'CLIENTES
        SQL = SQL & " and numserie >= 'A'"
        Tipo = 0
    Else
        If optVer(1).Value Then
            'Solo PAGO proveedores
            SQL = SQL & " and numserie = '2'"
            Tipo = 1
            
        Else
            SQL = SQL & " and numserie < '2'"
            Tipo = 2
        End If
    End If
    
    'La ordenacion
    SQL = SQL & " ORDER BY feccaja,numserie, numfactu,numorden"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        With miRsAux.Fields
            Set ItmX = ListView1(2).ListItems.Add
            ItmX.Text = Format(!feccaja, "dd/mm/yyyy")
            If Tipo < 2 Then
                ItmX.SubItems(1) = !codmacta
                ItmX.SubItems(2) = !Nommacta
            Else
                If !NUmSerie = "0" Then
                    ItmX.SubItems(1) = "SUMINIS."
                Else
                    ItmX.SubItems(1) = "TRASPASO"
                End If
                ItmX.SubItems(2) = !Ampliacion
            End If
            
            'Eecto
            Select Case Tipo
            Case 0
                ItmX.SubItems(3) = !NUmSerie & "/" & !numfactu & " - " & !numorden

            Case 1
                ItmX.SubItems(3) = !numfactu & " - " & !numorden
            '    ItmX.Tag = !codmacta & "|" & "|" & !numfactu & "|" & !numorden & "|" & !fecfactu & "|"
            Case Else
            '    ItmX.Tag = !codmacta & "|" & !numfactu & "|" & !numorden & "|" & !fecfactu & "|" & !NUmSerie & "|"
            End Select
            ItmX.SubItems(4) = Format(!impefect, FormatoImporte)
            ItmX.Tag = !codmacta & "|" & !NUmSerie & "|" & !numfactu & "|" & !numorden & "|" & !fecfactu & "|"
        End With
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub




Private Function HacerCierreCaja() As Boolean
Dim vSQL As String
Dim CtaPendAplicar As String
    On Error GoTo Ehacercierrecaja
    HacerCierreCaja = False
    
    
    CtaPendAplicar = DevuelveDesdeBD("par_pen_apli", "paramtesor", "codigo", "1", "N")
    If CtaPendAplicar = "" Then
        MsgBox "La cuenta para las partidas pendientes de aplicacion esta vacia o falta por configurar", vbExclamation
        Exit Function
    End If
    
    
    'Compruebo que hay datos para cerrar
    Set miRsAux = New ADODB.Recordset
    vSQL = "Select count(*) from scaja where ctacaja='" & LaCuentaDeCaja & "' and feccaja <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    miRsAux.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then vSQL = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If vSQL <> "" Then
        MsgBox "Ningun registro devuelto", vbExclamation
        Exit Function
    End If
    
    
    
    
    
    
    
    'Si hay k imprimir
    If Me.ChkCierre.Value Then
        vSQL = " and ctacaja='" & LaCuentaDeCaja & "' and feccaja <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        If ImpirmirListadoCaja(vSQL, True) Then
            With frmImprimir
                .OtrosParametros = ""
                .NumeroParametros = 0
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = 12
                .Show vbModal
            End With
        Else
            Exit Function
        End If
            
    End If
    
    
    
    'Hacemos el cierre
    Conn.BeginTrans
    If ContabilizarCierreCaja(CDate(txtFecha(1).Text), LaCuentaDeCaja, CtaPendAplicar) Then
        Conn.CommitTrans
    Else
        'Conn.RollbackTrans
        TirarAtrasTransaccion
    End If

    Exit Function
Ehacercierrecaja:
    MuestraError Err.Number, Err.Description
    
End Function
