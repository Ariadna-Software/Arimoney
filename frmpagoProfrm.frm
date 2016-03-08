VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagoPro 
   Caption         =   "Pago proveedores"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "frmpagoProfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   19
      Left            =   120
      MaxLength       =   36
      TabIndex        =   13
      Tag             =   "Iban|T|S|||spagop|iban|||"
      Text            =   "Text1"
      Top             =   4200
      Width           =   765
   End
   Begin VB.Frame frameContene 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4080
      TabIndex        =   54
      Top             =   1440
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "Doc. emitido"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Tag             =   "Fecha vencimiento|N|S|||spagop|emitdocum|||"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   55
         Tag             =   "Fecha vencimiento|N|S|||spagop|contdocu|||"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   5760
      MaxLength       =   15
      TabIndex        =   19
      Tag             =   "Referencia|T|S|||spagop|referencia|||"
      Text            =   "Text1"
      Top             =   4200
      Width           =   2325
   End
   Begin VB.Frame FrameEstaEnCaja 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   615
      Left            =   6960
      TabIndex        =   52
      Top             =   1320
      Width           =   1335
      Begin VB.CheckBox Check1 
         Caption         =   "Esta en caja"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Tag             =   "s|N|S|||spagop|estacaja|||"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   18
      Tag             =   "Transfer|N|S|||spagop|transfer|0000||"
      Text            =   "Text1"
      Top             =   4200
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   17
      Tag             =   "cuenta bancaria|T|S|||spagop|cuentaba|0000000000||"
      Text            =   "9999999999"
      Top             =   4200
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   15
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   16
      Tag             =   "entidad|T|S|||spagop|CC|00||"
      Text            =   "99"
      Top             =   4200
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   14
      Left            =   1560
      MaxLength       =   36
      TabIndex        =   15
      Tag             =   "entidad|T|S|||spagop|oficina|0000||"
      Text            =   "9999"
      Top             =   4200
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   960
      MaxLength       =   36
      TabIndex        =   14
      Tag             =   "entidad|T|S|||spagop|entidad|0000||"
      Text            =   "9999"
      Top             =   4200
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   12
      Left            =   120
      MaxLength       =   60
      TabIndex        =   21
      Tag             =   "Cta prevista|T|S|||spagop|text2csb|||"
      Text            =   "Text1"
      Top             =   5640
      Width           =   6405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   120
      MaxLength       =   80
      TabIndex        =   20
      Tag             =   "Cta prevista|T|S|||spagop|text1csb|||"
      Text            =   "123456789012345678901234567890123456"
      Top             =   4920
      Width           =   8085
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   4320
      TabIndex        =   12
      Tag             =   "Cta real pago|T|S|||spagop|ctabanc2|||"
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   3
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   3480
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Tag             =   "Cta prevista|T|N|||spagop|ctabanc1|||"
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "Text2"
      Top             =   3480
      Width           =   2835
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ultimo pago"
      Height          =   1095
      Left            =   3720
      TabIndex        =   38
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "Importe|N|S|||spagop|imppagad|#,##0.00||"
         Text            =   "1.999.999.00"
         Top             =   480
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   7
         Left            =   960
         TabIndex        =   9
         Tag             =   "Fecha vencimiento|F|S|||spagop|fecultpa|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1560
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   195
         Index           =   8
         Left            =   2640
         TabIndex        =   40
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   39
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Importe|N|N|||spagop|impefect|#,##0.00||"
      Text            =   "1.999.999.00"
      Top             =   2520
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Tag             =   "Fecha vencimiento|F|N|||spagop|fecefect|dd/mm/yyyy||"
      Text            =   "99/99/9999"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Tag             =   "Forma Pago|N|N|0||spagop|codforpa|||"
      Text            =   "Text1"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   1560
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Tag             =   "Cta. Cta proveedor|T|N|||spagop|ctaprove||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   840
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   5760
      TabIndex        =   2
      Tag             =   "Fecha Factura|F|N|||spagop|fecfactu|dd/mm/yyyy|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   7200
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Nº Vencimiento|N|N|0||spagop|numorden||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   885
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   6195
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Nº Factura|T|N|||spagop|numfactu||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   4095
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   6195
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Top             =   6195
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   960
      Top             =   6120
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
      TabIndex        =   30
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
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
            Object.ToolTipText     =   "Efectuar pago"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
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
         Left            =   6960
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Referencia"
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   53
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Transf/Pag dom"
      Height          =   255
      Index           =   17
      Left            =   4320
      TabIndex        =   51
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta"
      Height          =   255
      Index           =   16
      Left            =   2520
      TabIndex        =   50
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "D.C"
      Height          =   255
      Index           =   15
      Left            =   2160
      TabIndex        =   49
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Oficina"
      Height          =   195
      Index           =   14
      Left            =   960
      TabIndex        =   48
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "IBAN"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   47
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Texto 2 CSB34"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   46
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Texto 1 CSB34"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   45
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   3
      Left            =   5640
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta. real de pago"
      Height          =   195
      Index           =   10
      Left            =   4320
      TabIndex        =   44
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   2
      Left            =   1560
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta. prevista pago"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   42
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   5
      Left            =   840
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
      Height          =   195
      Index           =   6
      Left            =   1560
      TabIndex        =   37
      Top             =   2310
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "F. Vto"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   2280
      Width           =   795
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   1
      Left            =   1320
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de pago"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   6600
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   0
      Left            =   1200
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta. proveedor"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Fra"
      Height          =   195
      Index           =   4
      Left            =   5760
      TabIndex        =   29
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Vencimiento"
      Height          =   195
      Index           =   2
      Left            =   7080
      TabIndex        =   28
      Top             =   630
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Nº  Factura"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   27
      Top             =   600
      Width           =   1095
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
Attribute VB_Name = "frmPagoPro"
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
Private WithEvents frmF As frmFormaPago
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
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
Private NecesitaCuenta As String

Private TipoDePago As String

Dim BuscaChekc As String


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
                'Hacemos insertar
                'If ModificaDesdeFormulario(Me) Then
                    
                If ModificaDesdeFormularioClaves(Me, Cad) Then
                    'TerminaBloquear
                    DesBloqueaRegistroForm Me.Text1(0)
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
    DesBloqueaRegistroForm Me.Text1(0)
    'TerminaBloquear
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
        SituarData1 = False
        With Data1
            'Actualizamos el recordset
            .Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            .Recordset.MoveFirst
            While Not .Recordset.EOF
                If .Recordset!ctaprove = Text1(4).Text Then
                    If .Recordset!numfactu = Text1(1).Text Then
                        If Format(.Recordset!fecfactu, "dd/mm/yyyy") = Text1(2).Text Then
                            If CStr(.Recordset!numorden) = Text1(3).Text Then
                                SituarData1 = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
                .Recordset.MoveNext
            Wend
        End With
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
    NecesitaCuenta = ""
    'Añadiremos el boton de aceptar y demas objetos para insertar
    'cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    
    '###A mano
    Text1(4).SetFocus
    
    '## No puede poner a mano la transferencia
    Text1(17).Locked = True
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
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
'    Text1(0).Locked = True
'    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    '## No puede poner a mano la transferencia
    Text1(17).Locked = True
    'Para el campo
    If Text1(13).Text <> "" Then
        NecesitaCuenta = "1"
    Else
        NecesitaCuenta = ""
    End If
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'Comprobamos si se puede eliminar
    If Not SePuedeEliminar Then Exit Sub
    If PertenceAlgunoDocumentoEmitido Then Exit Sub
    
    
    '### a mano
    Cad = "Seguro que desea eliminar el registro:"
    Cad = Cad & vbCrLf & "Cta: " & Data1.Recordset.Fields(0) & " - " & Text2(0).Text
    Cad = Cad & vbCrLf & "Vencimiento: " & Data1.Recordset.Fields(1) & " - " & Data1.Recordset.Fields(3)
    Cad = Cad & vbCrLf & "Fecha VTO.: " & Data1.Recordset.Fields(5)
    Cad = Cad & vbCrLf & "importe: " & Data1.Recordset.Fields(6)
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
Dim Cad As String
Dim impo As Currency

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    
    If Not SePuedeEliminar Then Exit Sub
    
    
    
    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    If impo < 0 Then
        MsgBox "Los abonos no se realizan por caja", vbExclamation
        Exit Sub
    End If

    'Menos ya pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)
    

    
    If impo = 0 Then
        'YA ESTA TOTALMENTE PAGADO
        MsgBox "Totalmente pagado", vbExclamation
        Exit Sub
    End If
    
    
    
    'Devolvera muuuuchas cosas
    'serie factura fecfac numvto
    Cad = Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|"
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
        .Buttons(10).Image = 27
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Me.imgCuentas, 1, "Buscar cuenta"
    CargaImagenesAyudas Me.imgFecha, 2
    LimpiarCampos
    
    '## A mano
    NombreTabla = "spagop"
    Ordenacion = " ORDER BY ctaprove,numfactu,fecfactu,numorden"
        
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
    End If

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    lblIndicador.Caption = ""
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim Cad As String
    If CadenaDevuelta <> "" Then
        
        If DevfrmCCtas <> "" Then
    
            HaDevueltoDatos = True
            DevfrmCCtas = CadenaDevuelta
            
        Else
'                HaDevueltoDatos = True
'                Screen.MousePointer = vbHourglass
'                'Sabemos que campos son los que nos devuelve
'                'Creamos una cadena consulta y ponemos los datos
'                DevfrmCCtas = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
'                Cad = DevfrmCCtas
'                DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
'                Cad = Cad & " AND " & DevfrmCCtas
'                DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
'                Cad = Cad & " AND " & DevfrmCCtas
'                DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
'                Cad = Cad & " AND " & DevfrmCCtas
'                DevfrmCCtas = Cad
'                If DevfrmCCtas = "" Then Exit Sub
'                '   Como la clave principal es unica, con poner el sql apuntando
'                '   al valor devuelto sobre la clave ppal es suficiente
'                'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
'                'If CadB <> "" Then CadB = CadB & " AND "
'                'CadB = CadB & Aux
'                'Se muestran en el mismo form
'                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & DevfrmCCtas & " " & Ordenacion
'                PonerCadenaBusqueda
'                Screen.MousePointer = vbDefault
        End If
    Else
        DevfrmCCtas = ""
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(imgFecha(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0) = RecuperaValor(CadenaSeleccion, 1)
    Text2(1) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim Cad As String
    Screen.MousePointer = vbHourglass
    If Index = 1 Then
'        DevfrmCCtas = "-1"
'        Cad = "Código|codforpa|N|30·"
'        Cad = Cad & "Descripción|nomforpa|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = "sforpa"
'        frmB.vSQL = ""
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Formas de pago"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'            If DevfrmCCtas <> "-1" Then
'                Text1(0) = RecuperaValor(DevfrmCCtas, 1)
'                Text2(1) = RecuperaValor(DevfrmCCtas, 2)
'            End If
'        End If
    
        Set frmF = New frmFormaPago
        frmF.DatosADevolverBusqueda = "0|"
        frmF.Show vbModal
        Set frmF = Nothing
    
    Else
        'Cuentas
        Set frmCCtas = New frmColCtas
        DevfrmCCtas = ""
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        If DevfrmCCtas <> "" Then
            If Index > 0 Then
                Text1(7 + Index) = RecuperaValor(DevfrmCCtas, 1)
            Else
                Text1(4 + Index) = RecuperaValor(DevfrmCCtas, 1)
            End If
            Text2(Index).Text = RecuperaValor(DevfrmCCtas, 2)
        End If
    End If
    
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim F As Date
    If Text1(Index).Locked Then Exit Sub
    Set frmC = New frmCal
    F = Now
    If Text1(Index).Text <> "" Then
        If IsDate(Text1(Index).Text) Then F = CDate(Text1(Index).Text)
    End If
    frmC.Fecha = F
    imgFecha(2).Tag = Index
    frmC.Show vbModal
    
    Set frmC = Nothing
    
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


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo = 1 Then
        'BUSQUEDA
        If KeyCode = 112 Then HacerF1
    ElseIf Modo = 0 Then
        If KeyCode = 27 Then Unload Me
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
    
'    I = 0
'    If Modo < 3 Then
'        If Index = 4 Or Index = 9 Or Index = 10 Or Index = 0 Then I = 1
'    End If
'    If I = 0 Then Exit Sub
    
    'Campo con valor
    Select Case Index
    Case 4, 9, 10
            'Cuentas          'Cuentas
            'Cuentas          'Cuentas
        I = DevuelveText2Relacionado(Index)
        DevfrmCCtas = Text1(Index).Text
        If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
            Text1(Index).Text = DevfrmCCtas
            If Modo > 1 Then Text2(I).Text = SQL
            If Index = 4 Then PonerDatosBanco
        Else
            If Modo > 2 Then
                MsgBox SQL, vbExclamation
                Text1(Index).Text = ""
                Text2(I).Text = ""
                Ponerfoco Text1(Index)
            Else
                If DevfrmCCtas <> "" Then Text1(Index).Text = DevfrmCCtas
            End If
        End If
        
        'Poner la cuenta bancaria a partir de la cuenta
        If DevfrmCCtas <> "" Then
            If Modo > 2 And Index = 4 Then
                SQL = ""
                For I = 1 To 4
                    SQL = SQL & Text1(12 + I).Text
                Next I
                Valor = DevuelveLaCtaBanco(DevfrmCCtas)
                
               
                
                If Modo = 4 Then
                    If SQL <> "" Then
                        'Tienen puesta una cuenta
                        If Len(Valor) = 5 Then
                            SQL = "El proveedor no tienen cuenta bancaria. Quiere quitar la que estaba?"
                        Else
                            SQL = "Poner la cuenta bancaria del proveedor " & CStr(Valor) & "?"
                        End If
                        
                        If Text1(Index).Text = "" Then
                            SQL = "OK"
                            Valor = "|||||"
                        Else
                            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then SQL = ""
                        End If
                    End If
                Else
                    SQL = ""
                End If
                
                If SQL = "" Then
                       'SQL = DevuelveLaCtaBanco(DevfrmCCtas)
                       SQL = Valor
                       For I = 1 To 4
                           Text1(12 + I).Text = RecuperaValor(SQL, I)
                       Next I
                       Text1(19).Text = RecuperaValor(SQL, I) 'IBAN
                   
                End If

            End If
        End If
        
        
        
        
        
     Case 0
        'FORMA DE PAGO
        I = 1
        DevfrmCCtas = ""
        If Not IsNumeric(Text1(Index).Text) Then
            SQL = "Campo Forma pago debe ser numérico: " & Text1(Index).Text
            If Modo > 1 Then MsgBox SQL, vbExclamation
        Else
            SQL = "tipforpa"
            NecesitaCuenta = ""
            Valor = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", Text1(Index).Text, "N", SQL)
            If Valor = "" Then
                Valor = "Forma de pago inexistente: " & Text1(Index).Text
                If Modo > 1 Then MsgBox Valor, vbExclamation
            Else
                DevfrmCCtas = Valor
            End If
        End If
        Text2(I).Text = DevfrmCCtas
        If DevfrmCCtas = "" Then
            If Modo > 1 Then
                Text1(Index).Text = ""
                Ponerfoco Text1(Index)
            End If
        Else
            If SQL = 1 Then NecesitaCuenta = "1"
            PonerDatosBanco
            
            
        End If

        TipoDePago = SQL
        
    Case 2, 5, 7
        'FECHAS
        If Not EsFechaOK(Text1(Index)) Then
            If Modo > 2 Then
                MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                Ponerfoco Text1(Index)
            End If
        End If
        
    Case 6, 8
        'IMPORTES
        If Modo <= 2 Then Exit Sub
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
    Case 13 To 16
        If Index = 15 Then
            I = 2
            If Text1(15).Text = "**" Then Exit Sub
        Else
            If Index = 16 Then
                I = 10
            Else
                I = 4
            End If
        End If
          
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
        End If
        
        Text1(Index).Text = Right("0000000000" & Trim(Text1(Index).Text), I)
      
        
        SQL = ""
        For I = 13 To 16
            SQL = SQL & Text1(I).Text
        Next
        
        If Len(SQL) = 20 Then
            'OK. Calculamos el IBAN
            
            
            If Text1(19).Text = "" Then
                'NO ha puesto IBAN
                If DevuelveIBAN2("ES", SQL, SQL) Then Text1(19).Text = "ES" & SQL
            Else
                Valor = CStr(Mid(Text1(19).Text, 1, 2))
                If DevuelveIBAN2(CStr(Valor), SQL, SQL) Then
                    If Mid(Text1(19).Text, 3) <> SQL Then
                        
                        MsgBox "Codigo IBAN distinto del calculado [" & Valor & SQL & "]", vbExclamation
                        'Text1(49).Text = "ES" & SQL
                    End If
                End If
            End If
        End If

        
    Case 19
        Text1(Index).Text = UCase(Text1(Index).Text)
    End Select
            
End Sub

Private Sub PonerDatosBanco()
Dim SQL As String
    If Modo < 3 Then Exit Sub
    'Si ya tiene datos baconx, no hacemos nada
    If Text1(13).Text <> "" Then Exit Sub
    If Text1(4).Text <> "" And NecesitaCuenta <> "" Then
        SQL = "Select entidad,oficina,CC,cuentaba from cuentas where codmacta='" & Text1(4).Text & "'"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            For NumRegElim = 0 To 3
                Text1(NumRegElim + 13).Text = DBLet(miRsAux.Fields(NumRegElim), "T")
            Next NumRegElim
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    Else
            For NumRegElim = 0 To 3
                Text1(NumRegElim + 13).Text = ""
            Next NumRegElim
    
    End If
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
        frmVerCobrosPagos.Regresar = True
        frmVerCobrosPagos.OrdenarEfecto = False
        frmVerCobrosPagos.Cobros = False
        frmVerCobrosPagos.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            PonerDatoDevuelto CadenaDesdeOtroForm
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
        End If
        
        
        
End Sub

Private Sub PonerDatoDevuelto(CadenaDevuelta As String)
Dim Cad As String
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
    
    
    BuscaChekc = ""
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
        'cmdCancelar.Cancel = False
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
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = B
        If B Then
            Text1(I).BackColor = &H80000018
        ElseIf Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    frameContene.Enabled = Not B
    For I = 0 To 3
        imgCuentas(I).Visible = Not B
        
        'Me.imgFecha(I).Enabled = Not B
    Next I
    Me.imgFecha(2).Visible = Not B
    Me.imgFecha(5).Visible = Not B
    Me.imgFecha(7).Visible = Not B

    
        
    FrameEstaEnCaja.Enabled = (Modo = 1)
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    'AHora comprobare Si tiene cuenta bancaria, si es correcta
    
    DevfrmCCtas = ""
    
    For NumRegElim = 13 To 16
        If NumRegElim <> 15 Then DevfrmCCtas = DevfrmCCtas & Trim(Text1(NumRegElim).Text)
    Next NumRegElim
    If DevfrmCCtas <> "" Then
        If Len(DevfrmCCtas) <> 18 Then
            MsgBox "Longitud de cuenta bancaria incorrecta", vbExclamation
            Exit Function
        Else
            Me.Tag = CodigoDeControl(DevfrmCCtas)
            If Me.Tag <> Text1(15).Text Then
                Me.Tag = " CC funcion: " & Me.Tag & vbCrLf
                Me.Tag = "El CC no corresponde a la cuenta bancaria indicada." & vbCrLf & Me.Tag
                Me.Tag = Me.Tag & " CC escrito : " & Text1(15).Text & vbCrLf & vbCrLf & "¿Continuar?"
                
                If MsgBox(Me.Tag, vbQuestion + vbYesNo) = vbNo Then Exit Function
                
                DevfrmCCtas = Mid(DevfrmCCtas, 1, 8) & Me.Text1(15).Text & Mid(DevfrmCCtas, 9)
                Me.Tag = ""
                If Me.Text1(19).Text <> "" Then Me.Tag = Mid(Text1(19).Text, 1, 2)
                    
                If DevuelveIBAN2(CStr(Me.Tag), DevfrmCCtas, DevfrmCCtas) Then
                    If Me.Text1(19).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(19).Text = DevfrmCCtas
                    Else
                        If Mid(Text1(19).Text, 3) <> DevfrmCCtas Then
                            DevfrmCCtas = "Calculado : " & BuscaChekc & DevfrmCCtas
                            DevfrmCCtas = "Introducido: " & Me.Text1(19).Text & vbCrLf & DevfrmCCtas & vbCrLf
                            DevfrmCCtas = "Error en codigo IBAN" & vbCrLf & DevfrmCCtas & "Continuar?"
                            If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
                        End If
                    End If
                End If
                
                
                
                
            End If
        End If
    End If
    
    
    If CuentaBloqeada(Me.Text1(4).Text, CDate(Text1(2).Text), True) Then Exit Function
        
    
    DatosOk = B
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

    If PertenceAlgunoDocumentoEmitido Then Exit Sub

    If BloqueaRegistroForm(Me) Then BotonModificar
    'If BLOQUEADesdeFormulario(Me) Then BotonModificar
Case 8
    BotonEliminar
Case 10
    If Me.Data1.Recordset.EOF Then Exit Sub
    
    If Modo <> 2 Then Exit Sub
    
    If Not SePuedeEliminar Then Exit Sub
    
    If PertenceAlgunoDocumentoEmitido Then Exit Sub
        

    If BloqueaRegistroForm(Me) Then
        RealizarPagoCuenta
        DesBloqueaRegistroForm Text1(0)
    End If


Case 12
    mnSalir_Click
Case 14 To 17
    Desplazamiento (Button.Index - 14)
'Case 20
'    'Listado en crystal report
'    Screen.MousePointer = vbHourglass
'    CR1.Connect = Conn
'    CR1.ReportFileName = App.Path & "\Informes\list_Inc.rpt"
'    CR1.WindowTitle = "Listado incidencias."
'    CR1.WindowState = crptMaximized
'    CR1.Action = 1
'    Screen.MousePointer = vbDefault

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
    Text1_LostFocus 0
    Text1_LostFocus 9
    Text1_LostFocus 10
Exit Sub
EPonerCtasIVA:
    MuestraError Err.Number, "Poniendo valores ctas. IVA", Err.Description
End Sub


Private Function PertenceAlgunoDocumentoEmitido() As Boolean

    PertenceAlgunoDocumentoEmitido = False
    If Val(Data1.Recordset!emitdocum) = 1 Then
        If MsgBox("Pertence a un documento emtitido.  No deberia seguir con el proceso." & vbCrLf & vbCrLf & "Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then PertenceAlgunoDocumentoEmitido = True
    End If

End Function



Private Sub Ponerfoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SePuedeEliminar() As Boolean
    SePuedeEliminar = False

    If Not IsNull(Me.Data1.Recordset!transfer) Then
        If Val(Me.Data1.Recordset!transfer) > 0 Then
            MsgBox "Pertenece a una transferencia.", vbExclamation
            Exit Function
        End If
    End If
    
    If Val(DBLet(Data1.Recordset!estacaja)) > 0 Then
        MsgBox "Esta en caja", vbExclamation
        Exit Function
    End If
    
    SePuedeEliminar = True

End Function






Private Sub RealizarPagoCuenta()
Dim impo As Currency


  

    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    
    'Pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)
    
    'Si impo>0 entonces TODAVIA puedn pagarme algo
    If impo = 0 Then
        'Cosa rara. Esta todo el importe pagado
        Exit Sub
    End If
        
    frmParciales.Cobro = False 'PAGO
    frmParciales.Vto = Text1(4).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|"
    frmParciales.Importes = Text1(6).Text & "|" & Text1(8).Text & "|"
    frmParciales.Cta = Text1(4).Text & "|" & Text2(0).Text & "|" & Text1(9).Text & "|" & Text2(2).Text & "|"
    frmParciales.FormaPago = Val(TipoDePago)
    frmParciales.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'Hay que refrescar los datos
        lblIndicador.Caption = ""
        If SituarData1 Then
            
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
    If C = "" Then Text1(1).Text = "*"  'Para que busqu toooodo
    cmdAceptar_Click
End Sub
