VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRecpcionDoc 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmRecpcionDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSuma 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Llevado banco"
      Height          =   255
      Left            =   6240
      TabIndex        =   33
      Tag             =   "C|N|N|||scarecepdoc|LlevadoBanco|||"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Contab"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Tag             =   "C|N|N|||scarecepdoc|Contabilizada|||"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Fecha vencimiento|F|N|||scarecepdoc|fechavto|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   4920
      TabIndex        =   8
      Tag             =   "Importe|N|N|0||scarecepdoc|importe|#,##0.00||"
      Text            =   "Text1"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   3720
      TabIndex        =   6
      Tag             =   "Banco|T|N|||scarecepdoc|banco|||"
      Text            =   "Text1"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRecpcionDoc.frx":000C
      Left            =   1080
      List            =   "frmRecpcionDoc.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Talon|N|N|0||scarecepdoc|talon|||"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   0
      Tag             =   "Codigo|N|S|0||scarecepdoc|codigo||S|"
      Text            =   "Text1"
      Top             =   810
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text4"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   240
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Cliente|T|N|||scarecepdoc|codmacta|||"
      Text            =   "commor"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   2640
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Fecha recepcion|F|N|||scarecepdoc|fecharec|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   16
      Top             =   6360
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   7320
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   6360
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   11
      Top             =   6360
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   12
      Top             =   6360
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   5640
      TabIndex        =   13
      Top             =   6360
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6240
      Top             =   480
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   7320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Tag             =   "Referencia|T|N|||scarecepdoc|numeroref|||"
      Text            =   "Text1"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   7320
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRecpcionDoc.frx":0029
      Height          =   4455
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   7455
      _ExtentX        =   13150
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
      TabIndex        =   22
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
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
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Contabilizar recepcion"
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
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6720
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   495
      Left            =   4680
      Top             =   0
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
   Begin VB.Label lblSuma 
      Caption         =   "Suma"
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
      Left            =   7080
      TabIndex        =   35
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Vto."
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   32
      Top             =   600
      Width           =   930
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   1
      Left            =   4920
      Picture         =   "frmRecpcionDoc.frx":003E
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   31
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Banco"
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   30
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   29
      Top             =   540
      Width           =   975
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   2
      Left            =   840
      Picture         =   "frmRecpcionDoc.frx":00C9
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   3600
      Picture         =   "frmRecpcionDoc.frx":0ACB
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   195
      Index           =   9
      Left            =   1680
      TabIndex        =   28
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Id"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "F. Recep."
      Height          =   195
      Index           =   5
      Left            =   2640
      TabIndex        =   24
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "N� Referencia"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Menu mnOpcionesAsiPre 
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
Attribute VB_Name = "frmRecpcionDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const NO = "No encontrado"
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busquedaa
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'//////////////////////////////////
'//////////////////////////////////
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
Dim Ancho As Integer
'Dim colMes As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

Dim PrimeraVez As Boolean
Dim ImporteVto As Currency
Dim PosicionGrid As Integer





Private Sub PonerLineaModificadaSeleccionada()
    On Error GoTo E1


    Exit Sub
E1:
    Err.Clear
End Sub





Private Sub Check1_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub



Private Sub Check2_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Select Case Modo
    Case 1
        HacerBusqueda
    Case 3
        If Not DatosOk Then Exit Sub
        
        If InsertarRegistro Then
            NuevoTalonPagareDefecto False   'Para que memorize la ultima opcion
            If SituarData1(True) Then
                ModificandoLineas = 0
                
                PonerModo 5

                cmdCancelar.Caption = "Cabecera"
                lblIndicador.Caption = "Lineas detalle"
                
                AnyadirLinea
            End If
        End If
        
    Case 4
        If Not DatosOk Then Exit Sub
        
        If ModificaDesdeFormulario(Me) Then
            'Ha cambiado fecha vto
            CambiaFechaVto
            If SituarData1(False) Then
                lblIndicador.Caption = ""
                PonerModo 2
            Else
                PonerModo 0
            End If
        End If
      
    Case 5
        
        If Not AuxOK_ Then Exit Sub
        
        If InsertarModificar Then
            CargaGrid True
            'SIEMPRE SERA 1
            If ModificandoLineas = 1 Then
                ModificandoLineas = 0 'Para que vuelva a hacer el nuevo
                AnyadirLinea
            End If
        End If
    End Select
End Sub

Private Sub cmdAux_Click(Index As Integer)
Dim Im As Currency

        Im = 0
        If Me.txtSuma.Text <> "" Then Im = ImporteFormateado(txtSuma.Text)
        Im = ImporteFormateado(Text1(5).Text) - Im
        

        CadenaDesdeOtroForm = ""
        'Todos los cobros pendientes de este
        SQL = " scobro.codmacta = '" & Text1(2).Text & "' AND ( impcobro =0 or impcobro is null)"
        
        'MODIFICADO Agosto 2009
        SQL = " scobro.codmacta = '" & Text1(2).Text & "' AND estacaja=0 AND ( tiporem is null or tiporem>1)"
        
        'Docu recibido NO
        SQL = SQL & " AND recedocu = 0" 'por si acoaso
        
        frmVerCobrosPagos.ImporteGastosTarjeta_ = Im
        frmVerCobrosPagos.vSQL = SQL
        frmVerCobrosPagos.OrdenarEfecto = False
        frmVerCobrosPagos.Regresar = True
        frmVerCobrosPagos.Cobros = True
        frmVerCobrosPagos.DesdeRecepcionTalones = True 'Para que muestre el boton de dividir vencimiento
        frmVerCobrosPagos.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
                'Devuelve:  M|2700266|25/06/2008|1|
               SQL = " AND numserie = '" & RecuperaValor(CadenaDesdeOtroForm, 1) & "' AND codfaccl = " & RecuperaValor(CadenaDesdeOtroForm, 2)
               SQL = SQL & " and fecfaccl ='" & Format(RecuperaValor(CadenaDesdeOtroForm, 3), FormatoFecha) & "' AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 4)
               'El SQL es de todo el modulo
               PonerCamposVencimiento True
        End If
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
    
        'Si esta todavia sin llevar a banco haremos las sumas para ver si coinciden
        If Not ComprobarImportes Then Exit Sub
    
        CamposAux False, 0, False
        


        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
 
        
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If
        
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
        SQL = "Select * from scarecepdoc WHERE codigo =" & Text1(4).Text
        Data1.RecordSource = SQL
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(Data1.Recordset!Codigo) = Text1(4).Text Then
                lblIndicador.Caption = ""
                SituarData1 = True
                Exit Function
                
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
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
   ' CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
   ' PonerCadenaBusqueda True
    
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3

    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    

    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    NuevoTalonPagareDefecto True
    'Combo1.ListIndex = 1 'Talon
    
    
    Ponerfoco Combo1
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "B�squeda"
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
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    
    
    
 
    'A�adiremos el boton de aceptar y demas objetos para insertar
    cmdCancelar.Caption = "Cancelar"
    cmdAceptar.Caption = "&Modificar"
    PonerModo 4



    


    'Si tienen NO tiene lineas dejaremos modificar la cuenta contable
    If Adodc1.Recordset.EOF Then Text1(2).Enabled = True
    


    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    Ponerfoco Text1(0)
End Sub

Private Sub BotonEliminar()
Dim Importe As Currency
Dim Ok As Boolean
    If Data1.Recordset.EOF Then Exit Sub
    
    
    'Marzo 2010. Ya comrpueba en puederealizaraccion si tiene vtos asociados
'    If Not adodc1.Recordset.EOF Then
'        If Not (Check1.Value = 1 And Check2.Value = 1) Then
'            MsgBox "Elimine primero los vencimientos asociados al documento", vbExclamation
'            Exit Sub
'        End If
'    End If


    
    
    SQL = DevuelveDesdeBD("Contabilizada", "scarecepdoc", "codigo", Text1(4).Text)
    If SQL = "1" Then
        'Esta realizado el apunte. Hay que deshacer
        SQL = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", Text1(4).Text)
        If SQL = "" Then SQL = "0"
        I = 0
        If CCur(SQL) <> ImporteFormateado(Text1(5).Text) Then
            SQL = CStr(CCur(SQL) - ImporteFormateado(Text1(5).Text))
            If CCur(SQL) > 0 Then
                I = -1   'Mayor las lineas que el importe del talon
            Else
                I = 1   'Mayor el total que la suma de las lineas
            End If
        End If
        
        If Not HacerDES_Contabilizacion_(I) Then Exit Sub
        
        
        
    
        
        
    Else
        'NO ha hecho nada, se borra directamente
        If MsgBox("�Desea eliminar el documento recibido?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    
    Conn.BeginTrans
    NumRegElim = Data1.Recordset.AbsolutePosition
    
        I = 0
        If Not Adodc1.Recordset Is Nothing Then
            If Not Adodc1.Recordset.EOF Then
               Adodc1.Recordset.MoveFirst
               While Not Adodc1.Recordset.EOF
               
        
               
                   'Obtengo el importe del vto
                   SQL = MontaSQLDelVto(False)
                   SQL = SQL & " AND 1 " 'Para hacer un truqiot
                   SQL = DevuelveDesdeBD("impcobro", "scobro", SQL, "1")
                   If SQL = "" Then SQL = "0"
                   Importe = CCur(SQL)
                   If Importe <> Adodc1.Recordset!Importe Then
                       'TODO EL IMPORTE estaba en la linea. Fecultco a NULL
                       I = 1
                       Importe = Importe - Adodc1.Recordset!Importe
                   Else
                       I = 0
                   End If
               
                   SQL = "UPDATE scobro SET recedocu=0,reftalonpag = NULL"
                   If I = 0 Then
                       SQL = SQL & ", impcobro = NULL, fecultco = NULL"
                   Else
                       SQL = SQL & ", impcobro = " & TransformaComasPuntos(CStr(Importe))  'NO somos capace sde ver cual fue la utlima fecha de amortizacion
                   End If
                   SQL = SQL & ", obs= NULL"
                   SQL = SQL & " WHERE " & MontaSQLDelVto(False)
                   
                   If Not EjecutarSQL(SQL) Then
                       MsgBox "Error actualizadno scobro", vbExclamation
                       I = 100
                       Adodc1.Recordset.MoveLast
                   End If
                   
                   Adodc1.Recordset.MoveNext
               Wend
            End If
        End If
        If I = 100 Then
            Ok = False
        Else
            Ok = Eliminar
        End If
    If Ok Then
        Conn.CommitTrans
        
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            CargaGrid False
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
    Else
        'Conn.RollbackTrans
        TirarAtrasTransaccion

    End If

    
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim AUX As String

'If Data1.Recordset.EOF Then
'    MsgBox "Ning�n registro devuelto.", vbExclamation
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









Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub Form_Activate()
Dim B As Boolean
  
    If PrimeraVez Then
        B = False
        PrimeraVez = False
        PonerModo 0
        CargaGrid False
        

        
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    LimpiarCampos
    PrimeraVez = True
    CadAncho = False

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        
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
    
    Caption = "Recepcion de documentos TALON,PAGARE (" & vEmpresa.nomresum & ")"
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
       ' Me.Width = 12000
       ' Me.Height = Screen.Height
    End If
    Me.Height = 8625
    'Los campos auxiliares
    CamposAux False, 0, True
    


    '## A mano
    NombreTabla = "scarecepdoc"
    Ordenacion = " ORDER BY codigo"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn



End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    Me.Combo1.ListIndex = -1
    Me.Check1.Value = 0
    Me.Check2.Value = 0
    txtSuma.Text = ""
    lblIndicador.Caption = ""
End Sub


'Private Sub Form_Resize()
'If Me.WindowState <> 0 Then Exit Sub
'If Me.Width < 11610 Then Me.Width = 11610
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim AUX As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        AUX = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
        CadB = AUX

        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda False
        Screen.MousePointer = vbDefault
    End If

End Sub









Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(I).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



'Private Sub Image1_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        'Cta contrapartida
'
'    Case 1
'
'    Case 2
'
'    Case 3
'
'    End Select
'End Sub

Private Sub imgppal_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub

    Select Case Index
    Case 0, 1
        'FECHA
        If Index = 0 Then
            I = 1
        Else
            I = 6
        End If
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(I).Text <> "" Then frmF.Fecha = CDate(Text1(I).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    
    Case 2
    
       ' If Text1(2).Enabled Then   'Solo insertando
            Set frmCCtas = New frmColCtas
            SQL = ""
            frmCCtas.DatosADevolverBusqueda = "0"
            frmCCtas.Show vbModal
            Set frmCCtas = Nothing
            If SQL <> "" Then
                Text1(2) = RecuperaValor(SQL, 1)
                Text5.Text = RecuperaValor(SQL, 2)
            End If
       ' End If
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
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    If Modo = 5 Then Exit Sub
    

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
Dim EntrarEnSelect As Boolean
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    'Si estamos insertando o modificando o buscando
    'If Index = 2 Then Stop
    EntrarEnSelect = False
    If Modo = 3 Or Modo = 4 Then
        EntrarEnSelect = True
    Else
        If Modo = 1 And Index = 2 Then EntrarEnSelect = True
    End If
    If EntrarEnSelect Then
        If Text1(Index).Text = "" Then
            If Index = 0 Then
               
            Else
                If Index = 2 Then Text5.Text = ""
            End If
            Exit Sub
        End If
        Select Case Index
        Case 0
'            'Tipo diario
'            If Not EsNumerico(Text1(Index).Text) Then
'                Text1(Index).Text = ""
'                PonerFoco Text1(Index)
'            End If


        
            
        Case 1, 6

            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
                Text1(Index).Text = ""
                Ponerfoco Text1(Index)
            End If
            
        Case 2
        
            RC = CByte(CuentaCorrectaUltimoNivelTXT(Text1(2), Text5))
            If RC = 0 Then
                'Error. En busqueda dejamos pasar
                
                If Modo <> 1 Then
                    MsgBox Text5.Text, vbExclamation
                    Text1(2).Text = ""
                    Ponerfoco Text1(2)
                End If
                Text5.Text = ""
            End If
        Case 5

            FormatTextImporte Text1(Index)
            If Text1(Index).Text = "" Then Ponerfoco Text1(Index)
        End Select
    End If
End Sub



Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(4), 7, "Cod:")
        Cad = Cad & "T|if(talon=0,""P"",""T"")|T|5�"
        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha Vto")
        Cad = Cad & ParaGrid(Text1(0), 25, "Referencia")
        Cad = Cad & ParaGrid(Text1(3), 14, "Banco")
        Cad = Cad & "Cta|cuentas.codmacta|T|12�"
        Cad = Cad & "Titulo|nommacta|T|22�"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla & ",cuentas"
            Cad = NombreTabla & ".codmacta =cuentas.codmacta"
            If CadB <> "" Then Cad = Cad & " AND " & CadB
            frmB.vSQL = Cad
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|"
            frmB.vTitulo = "Recepcion documentos"
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
        MsgBox "No hay ning�n registro en la tabla de recepcion de documentos", vbInformation
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
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    'Cargamos datos extras

    If Text1(2).Text = "" Then
        SQL = ""
    Else
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(2).Text, "N")
        If SQL = "" Then SQL = "Error en cuenta contable"
    End If
    Text5.Text = SQL
    PonerImporteLinea
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
        Text1(4).BackColor = &HFEF7E4
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nueva recepcion"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar recepcion"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar  recepcion"
    End If
    
 
        
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea  recepcion"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea  recepcion"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea  recepcion"
    End If
    PonerOpcionesMenuGeneral Me
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B

    'Modo 2. Hay datos y estamos visualizandolos
    B = (Modo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B
    Toolbar1.Buttons(11).Enabled = B

        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.Visible = B Or Modo = 1
    'PRueba###
    


    '
    B = B Or (Modo = 5)
    mnOpcionesAsiPre.Enabled = Not B
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
    Text1(2).Enabled = (Modo = 3 Or Modo = 1) 'Solo insertar
    B = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    
    Me.Check1.Enabled = (Modo = 1) Or (B And vUsu.Nivel = 0)
    Me.Check2.Enabled = Me.Check1.Enabled
    Combo1.Enabled = B
    
    Text1(0).Enabled = B
    Text1(1).Enabled = B
    Text1(3).Enabled = B
    Text1(5).Enabled = B
    Text1(6).Enabled = B
    
    If Modo = 4 Then
        'Esta contabilizada
        If Me.Check1.Value = 1 Then
            ' no dejaremos cambiar el importe tampoc
            Text1(5).Enabled = False
            Combo1.Enabled = False
        End If
    End If
    'El text
    B = (Modo = 2) Or (Modo = 5)
    
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'FALTA###
    Toolbar1.Buttons(7).Enabled = Modo = 2
    mnModificar.Enabled = Modo = 2
    
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B


   
   
    If Modo <= 2 Then
         Me.cmdAceptar.Caption = "Aceptar"
         Me.cmdCancelar.Caption = "Cancelar"
    End If
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui a�adiremos controles para datos especificos. Esto es, si hay imagenes en el form
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
    
    'Me.lblSuma.Visible = Modo = 5
    'txtSuma.Visible = Modo = 5
End Sub


Private Function DatosOk() As Boolean
    
    Dim B As Boolean
    B = CompForm(Me)
    

    DatosOk = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub HacerToolBar(Boton As Integer)


    Select Case Boton
    Case 1
        BotonBuscar
    Case 2
        BotonVerTodos
    Case 6
        If Modo <> 5 Then
            BotonAnyadir
        Else
            'A�ADIR linea factura
            AnyadirLinea
        End If
    Case 7
        If Modo <> 5 Then
            'Intentamos bloquear la cuenta
            If PuedeRealizarAccion(False, True, False) Then BotonModificar
        Else
            'MODIFICAR linea factura
            ModificarLinea
        End If
    Case 8
        If Modo <> 5 Then
            If PuedeRealizarAccion(False, False, True) Then BotonEliminar
        Else
           
            EliminarLineaFactura
        End If
    Case 10
    
        If Not PuedeRealizarAccion(False, False, False) Then Exit Sub
        PonerImporteLinea
        'Nuevo Modo
        PonerModo 5

        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
        
    Case 11
        'Contabilizar
        
        I = 1
        If Combo1.ListIndex = 0 Then
            'PAGARE. Ver si tiene cta puente pagare
            If vParam.PagaresCtaPuente Then I = 0
        Else
            If vParam.TalonesCtaPuente Then I = 0
        End If
        If I = 1 Then
            MsgBox "Falta configurar en parametros", vbExclamation
            Exit Sub
        End If
        
        If Not PuedeRealizarAccion(True, False, False) Then Exit Sub
        
        
        SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", Text1(4).Text)
        If SQL = "" Then SQL = "0"
        If Val(SQL) = 0 Then
            MsgBox "No tiene vencimientos asociados", vbExclamation
            Exit Sub
        End If
        
        
        
        
        'Los importes
        SQL = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", Text1(4).Text)
        If SQL = "" Then SQL = "0"
        I = 0
        If CCur(SQL) <> ImporteFormateado(Text1(5).Text) Then
            SQL = CStr(CCur(SQL) - ImporteFormateado(Text1(5).Text))
            If CCur(SQL) > 0 Then
                I = -1   'Mayor las lineas que el importe del talon
            Else
                I = 1   'Mayor el total que la suma de las lineas
            End If
                
            SQL = "Suma de importes distintos del importe del talon: " & SQL
            If vUsu.Nivel <= 1 Then
                SQL = SQL & vbCrLf & "Seguro que desea continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Else
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
        End If
        
        
        'Hacemos contabilizacion
        HacerContabilizacion I
        
    Case 13
        'Imprimir
        frmListado.Opcion = 24
        frmListado.Show vbModal
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
    DataGrid1.Columns(0).Caption = "Serie"
    DataGrid1.Columns(0).Width = 800
    
    DataGrid1.Columns(1).Caption = "Factura"
    DataGrid1.Columns(1).Width = 2395
    

    DataGrid1.Columns(2).Caption = "Fecha"
    DataGrid1.Columns(2).Width = 1105
    DataGrid1.Columns(2).NumberFormat = "dd/mm/yyyy"
    
    DataGrid1.Columns(3).Caption = "Vto"
    DataGrid1.Columns(3).Width = 800
    
    DataGrid1.Columns(4).Caption = "Importe"
    DataGrid1.Columns(4).Width = 1000
    DataGrid1.Columns(4).NumberFormat = FormatoImporte
    DataGrid1.Columns(4).Alignment = dbgRight
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        
        txtaux(0).Left = DataGrid1.Left + 330
        txtaux(0).Width = DataGrid1.Columns(0).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(1).Left + DataGrid1.Left - cmdAux(0).Width - 15
                
        anc = DataGrid1.Left + 15
        txtaux(1).Left = DataGrid1.Columns(1).Left + anc
        txtaux(1).Width = DataGrid1.Columns(1).Width - 45
    
        txtaux(2).Left = DataGrid1.Columns(2).Left + anc
        txtaux(2).Width = DataGrid1.Columns(2).Width - 45
    
        txtaux(3).Left = DataGrid1.Columns(3).Left + anc
        txtaux(3).Width = DataGrid1.Columns(3).Width - 45

        
        'Concepto
        txtaux(4).Left = DataGrid1.Columns(4).Left + anc
        txtaux(4).Width = DataGrid1.Columns(4).Width - 45
        

       
        CadAncho = True
    End If
        
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    DataGrid1.Tag = "Calculando"

    If Modo = 5 Then PonerImporteLinea
    

    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub PonerImporteLinea()
Dim C As String
        C = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", Text1(4).Text)
        If C = "" Then C = "0"
        txtSuma.Text = Format(C, FormatoImporte)

End Sub



Private Function MontaSQLCarga(Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Bas�ndose en la informaci�n proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim SQL As String
    SQL = "SELECT numserie,numfaccl,fecfaccl,numvenci,importe FROm slirecepdoc WHERE id = "
    If Enlaza Then
        SQL = SQL & Data1.Recordset!Codigo
    Else
        SQL = SQL & "-1"
    End If
    SQL = SQL & " ORDER BY numserie,numfaccl,fecfaccl,numvenci"
    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea()
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
   
    
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
    LLamaLineas anc, 1, True

    'Ponemos el foco
    Ponerfoco txtaux(0)
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
  
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtaux(0).Text = Adodc1.Recordset.Fields!codmacta
    txtaux(1).Text = Adodc1.Recordset.Fields!Nommacta
    txtaux(2).Text = DataGrid1.Columns(4).Text
    txtaux(3).Text = DataGrid1.Columns(5).Text
    txtaux(4).Text = DataGrid1.Columns(6).Text
    txtaux(5).Text = DataGrid1.Columns(8).Text
    Cad = DBLet(Adodc1.Recordset.Fields!timported)
    If Cad <> "" Then
        txtaux(6).Text = Format(Cad, "0.00")
    Else
        txtaux(6).Text = Cad
    End If
    Cad = DBLet(Adodc1.Recordset.Fields!timporteH)
    If Cad <> "" Then
        txtaux(7).Text = Format(Cad, "0.00")
    Else
        txtaux(7).Text = Cad
    End If
    txtaux(8).Text = DBLet(Adodc1.Recordset.Fields!codccost)


    LLamaLineas anc, 2, False
    Ponerfoco txtaux(0)
End Sub

Private Sub EliminarLineaFactura()
Dim Importe As Currency

    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
   
    SQL = "Va a eliminar la linea: "
    SQL = SQL & Adodc1.Recordset!NUmSerie & Adodc1.Recordset!numfaccl & "  / " & Adodc1.Recordset!numvenci
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        NumRegElim = Adodc1.Recordset.AbsolutePosition
        
        
        'Obtengo el importe del vto
        SQL = MontaSQLDelVto(False)
        SQL = SQL & " AND 1 " 'Para hacer un truqiot
        SQL = DevuelveDesdeBD("impcobro", "scobro", SQL, "1")
        If SQL = "" Then SQL = "0"
        Importe = CCur(SQL)
        If Importe <> Adodc1.Recordset!Importe Then
            'TODO EL IMPORTE estaba en la linea. Fecultco a NULL
            I = 1
            Importe = Importe - Adodc1.Recordset!Importe
        Else
            I = 0
        End If
        
        SQL = "Delete from slirecepdoc"
        SQL = SQL & " WHERE id =" & Data1.Recordset!Codigo
        SQL = SQL & " AND " & MontaSQLDelVto(False)
        SQL = Replace(SQL, "codfaccl", "numfaccl")   'Me paso por gilipolla
        SQL = Replace(SQL, "numorden", "numvenci")  'idem, por no llamar igual a los campos
        DataGrid1.Enabled = False
        Conn.Execute SQL
        
        'Updateo en scbro reestableciendo los valores
        
        SQL = "UPDATE scobro SET recedocu=0,reftalonpag = NULL"
        If I = 0 Then
            SQL = SQL & ", impcobro = NULL, fecultco = NULL"
        Else
            SQL = SQL & ", impcobro = " & TransformaComasPuntos(CStr(Importe))  'NO somos capace sde ver cual fue la utlima fecha de amortizacion
        End If
        SQL = SQL & ", obs = NULL"
        SQL = SQL & " WHERE " & MontaSQLDelVto(False)
        If Not EjecutarSQL(SQL) Then MsgBox "Error actualizadno scobro", vbExclamation
        
        CargaGrid True
        DataGrid1.Enabled = True
        PosicionaLineas CInt(NumRegElim)
    End If
End Sub

Private Sub PosicionaLineas(Pos As Integer)
    On Error GoTo EPosicionaLineas
    If Pos > 1 Then
        If Pos >= Adodc1.Recordset.RecordCount Then Pos = Adodc1.Recordset.RecordCount - 1
        Adodc1.Recordset.Move Pos
    End If
    
    Exit Sub
EPosicionaLineas:
    Err.Clear
End Sub




'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------




Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid DataGrid1
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)


    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B

    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim I As Integer
    
    
    DataGrid1.Enabled = Not Visible

    For I = 0 To txtaux.Count - 1
        txtaux(I).Visible = Visible
        txtaux(I).Top = Altura
    Next I
    
    cmdAux(0).Visible = Visible
    cmdAux(0).Top = Altura
    If Limpiar Then
        For I = 0 To txtaux.Count - 1
            txtaux(I).Text = ""
        Next I
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), Modo

End Sub

Private Sub txtaux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'Esto sera k hemos pulsado el ENTER
        txtAux_LostFocus Index
        'cmdAceptar_Click   FALTA###
    Else
        If KeyCode = 113 Then
            
            
        Else
            'Ha pulsado F5. Ponemos linea anterior
            Select Case KeyCode
            Case 116
               
                
            Case 117
                'F6

                
            Case Else
                If (Shift And vbCtrlMask) > 0 Then
                    If UCase(Chr(KeyCode)) = "B" Then
                        'OK. Ha pulsado Control + B
                        '----------------------------------------------------
                        '----------------------------------------------------
                        '
                        ' Dependiendo de index lanzaremos una opcion uotra
                        '
                        '----------------------------------------------------
                        
                        'De momento solo para el 5. Cliente
                        Select Case Index
                        Case 4
                            'txtaux(4).Text = ""
                           ' Image1_Click 1
                        Case 8
                            'txtaux(8).Text = ""
                            'Image1_Click 2
                        End Select
                     End If
                End If
            End Select
        End If
    End If
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
        If KeyAscii = 112 Then
        
        End If
        'If KeyAscii = 43 Then
        '    If PulsadoMas(Index, KeyAscii) Then KeyAscii = 0
        'End If
    End If
End Sub



Private Sub txtaux_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                LanzaPantalla Index
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    

        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
           ' Select Case Index
           ' Case 0
           '
           '
           ' Case 3
           '
           ' Case 4
           '
           ' End Select
            Exit Sub
        End If
        

        
        Select Case Index
        Case 0
            txtaux(0).Text = UCase(txtaux(0).Text)

        Case 1, 3
            If Not EsNumerico(txtaux(Index).Text) Then
                
                txtaux(Index).Text = ""
                Ponerfoco txtaux(Index)
            End If
        
          
                
        Case 2
            If Not EsFechaOK(txtaux(Index)) Then
                txtaux(Index).Text = ""
                Ponerfoco txtaux(Index)
            End If
            
        Case 4
            FormatTextImporte txtaux(Index)
            If txtaux(Index).Text = "" Then
                Ponerfoco txtaux(Index)
            Else
                'El importe no puede ser mayor
                If ImporteFormateado(txtaux(Index).Text) > ImporteVto Then
                    MsgBox "El importe NO puede ser mayor al del vencimiento", vbExclamation
                    Ponerfoco txtaux(Index)
                End If
            End If
        End Select
        
        
        If Index = 0 Or Index = 1 Then
            If txtaux(0).Text <> "" And txtaux(1).Text <> "" Then PonerCamposVencimiento False
        End If
End Sub



Private Sub PonerCamposVencimiento(DesdeElButon As Boolean)
Dim Cad As String
Dim Importe As Currency

        'Veresmos si existe un unico vto para esta factura
        Cad = "Select numserie,codfaccl,fecfaccl,numorden,impvenci,impcobro,tipforpa,Gastos from scobro,sforpa"
        Cad = Cad & "  WHERE scobro.codforpa=sforpa.codforpa"
        Cad = Cad & " AND codmacta ='" & Text1(2).Text & "'"
        'Numero de serie y numfac
        If Not DesdeElButon Then
            Cad = Cad & " AND numserie ='" & txtaux(0).Text & "' AND codfaccl = " & txtaux(1).Text
        Else
            Cad = Cad & SQL  'SQL traera los datos del venciemietno
        End If
            
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        SQL = ""
        While Not miRsAux.EOF
            Importe = miRsAux!impvenci + DBLet(miRsAux!Gastos, "N")
            Importe = Importe - DBLet(miRsAux!impcobro, "N")
            
            If Importe = 0 Then
                
                    If IsNull(miRsAux!impcobro) Then
                        Cad = "Importe es cero"
                    Else
                        Cad = "Totalmente cobrado"
                    End If
                    MsgBox Cad, vbExclamation
                
            Else
                NumRegElim = vbTalon
                If Me.Combo1.ListIndex = 0 Then NumRegElim = vbPagare
                
                    
                    
                
                SQL = SQL & miRsAux!NUmSerie & "|" & miRsAux!codfaccl & "|"
                SQL = SQL & miRsAux!numorden & "|" & Format(miRsAux!fecfaccl, "dd/mm/yyyy") & "|"
                SQL = SQL & Importe & "|" 'Importe
                'si la forma de pago corresponde al documento que estamos procesando
                SQL = SQL & Abs((miRsAux!tipforpa = NumRegElim)) & "|:"
                I = I + 1
                
            End If
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        If I > 0 Then
            'HAY DATOS
            If I = 1 Then
                'SOLO HAY UNO
                SQL = Mid(SQL, 1, Len(SQL) - 1) 'Le quito los dos puntos
                
                
                
            Else
                'Hay mas de uno. Mostraremos una windows
            
            
                SQL = ""
            End If
            'Pongo los datos
            If SQL <> "" Then PonerDatosVencimiento SQL, DesdeElButon
        End If
End Sub



Private Function AuxOK_() As Boolean
Dim Importe As Currency
Dim Cad As String
    AuxOK_ = False
    For I = 0 To txtaux.Count - 1
        If txtaux(I).Text = "" Then
            MsgBox "Campo obligatorio", vbExclamation
            Ponerfoco txtaux(I)
            Exit Function
        End If
    Next
    
    If ModificandoLineas = 1 Then
        SQL = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", Text1(4).Text)
        If SQL = "" Then SQL = "0"
        Importe = ImporteFormateado(SQL)
        If Importe + ImporteFormateado(txtaux(4).Text) > ImporteFormateado(Text1(5).Text) Then
            SQL = CStr(Importe + ImporteFormateado(txtaux(4).Text) - ImporteFormateado(Text1(5).Text))
            SQL = "Suma de importes execede del importe del talon : " & SQL & vbCrLf & vbCrLf
            Importe = ImporteFormateado(Text1(5).Text) - Importe
            SQL = SQL & "Importe maximo del vto: " & Importe
            SQL = SQL & vbCrLf & vbCrLf & "�Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
                'Pongo el foco en el campo
                Ponerfoco txtaux(4)
                Exit Function
            End If
        End If

    
        'Ahora veremos si esta introduciendo un VTO sin elimporte total....
        Cad = "Select impvenci,impcobro,Gastos from scobro"
        Cad = Cad & "  WHERE numserie ='" & txtaux(0).Text & "' AND codfaccl = " & txtaux(1).Text
        Cad = Cad & "  AND fecfaccl='" & Format(txtaux(2).Text, FormatoFecha) & "' AND numorden = " & txtaux(3).Text
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
                Importe = miRsAux!impvenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
                Importe = Importe - ImporteFormateado(txtaux(4).Text)
                If Importe > 0 Then
                    Cad = "Deberia dividir el vencimiento si no lo va a remesar por el total pendiente."
                    Cad = Cad & vbCrLf & vbCrLf & "�Continuar?"
                    If MsgBox(Cad, vbQuestion + vbYesNo + vbMsgBoxRight) = vbYes Then AuxOK_ = True
                Else
                    AuxOK_ = True
                End If
        Else
            MsgBox "Vencimiento NO encontrado.  Funcion: auxok", vbCritical
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If

    
End Function





Private Function InsertarModificar() As Boolean
Dim Importe As Currency

    On Error GoTo EInsertarModificar
    Set miRsAux = New ADODB.Recordset
    
    InsertarModificar = False
    
    'Cargaremos el VTO de la scobro
    SQL = MontaSQLDelVto(True)
    SQL = " WHERE scobro.codforpa=sforpa.codforpa AND " & SQL
    SQL = "select scobro.*,tipforpa from scobro,sforpa " & SQL
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "El vencimiento introducido no se corresponde con ning�n cobro pendiente", vbExclamation
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Function
    End If
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS

        SQL = "insert into `slirecepdoc` (`id`,`numserie`,`numfaccl`,`fecfaccl`, "
        SQL = SQL & "`numvenci`,`importe`,`contabilizado`) VALUES ("
        SQL = SQL & Data1.Recordset!Codigo & ",'"
        SQL = SQL & txtaux(0).Text & "',"
        SQL = SQL & txtaux(1).Text & ",'"
        
        SQL = SQL & Format((txtaux(2).Text), FormatoFecha) & "'," & txtaux(3).Text & ","
        SQL = SQL & TransformaComasPuntos(ImporteFormateado(txtaux(4).Text)) & ",0)"
    Else
    
        'MODIFICAR
'        SQL = "UPDATE linapu SET "
'        SQL = SQL & " codmacta = '" & txtAux(0).Text & "',"
'        SQL = SQL & " numdocum = '" & DevNombreSQL(txtAux(2).Text) & "',"
'        SQL = SQL & " codconce = " & txtAux(4).Text & ","
'        SQL = SQL & " ampconce = '" & DevNombreSQL(txtAux(5).Text) & "',"
''        If txtaux(6).Text = "" Then
'          SQL = SQL & " timporteD = " & ValorNulo & "," & " timporteH = " & TransformaComasPuntos(txtaux(7).Text) & ","
'        SQL = SQL & " WHERE linapu.linliapu = " & Linliapu
'        SQL = SQL & " AND linapu.numdiari=" & Data1.Recordset!numdiari
'        SQL = SQL & " AND linapu.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
'        SQL = SQL & "' AND linapu.numasien=" & Data1.Recordset!Numasien & ";"
'
    End If
    Conn.Execute SQL
    
    
    
    
    
    
    
    
    
    'Segunda parte del meollo. En la scobro MARCAREMOS el vencimiento
    '
    '      Documento recibido
    '      Importe cobrado
    '      si no tiene forma de pago talon / pager se la pongo

    
    SQL = "UPDATE scobro SET recedocu=1,reftalonpag = '" & DevNombreSQL(Text1(0).Text) & "'"
    Importe = DBLet(miRsAux!impcobro, "N") + ImporteFormateado(txtaux(4).Text)
    SQL = SQL & ", impcobro = " & TransformaComasPuntos(CStr(Importe))
    SQL = SQL & ", fecultco = '" & Format(Text1(1).Text, FormatoFecha) & "'"
    'Febrero 2010
    'Fecha vencimiento tb le pongo la de la recpcion
    SQL = SQL & ", fecvenci = '" & Format(Text1(6).Text, FormatoFecha) & "'"
    'BANCO LO PONGO EN OBSERVACION
    SQL = SQL & ", obs = '" & DevNombreSQL(Text1(3).Text) & "'"
    'Si no era forma de pago talon/pagare la pongo
    If Me.Combo1.ListIndex = 0 Then
        I = vbPagare
    Else
        I = vbTalon
    End If
    If miRsAux!tipforpa <> I Then
        'AQUI BUSCARE una forma de pago
        I = Val(DevuelveDesdeBD("codforpa", "sforpa", "tipforpa", CStr(I)))
        If I > 0 Then SQL = SQL & ", codforpa = " & I
        
    End If
    SQL = SQL & " WHERE " & MontaSQLDelVto(True)
    miRsAux.Close
    
    If Not EjecutarSQL(SQL) Then MsgBox "Actualizando scobro. Avise soporte", vbExclamation
    
    InsertarModificar = True
    
EInsertarModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
    Set miRsAux = Nothing
End Function
 



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
        'Alguna comprobacion
        
        
        
        'Lineas
        Conn.Execute "Delete  from slirecepdoc WHERE id =" & Text1(4).Text
        
        'Cabeceras
        Conn.Execute "Delete  from scarecepdoc WHERE codigo =" & Text1(4).Text
        
                

        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        
        
        Eliminar = False
    Else
       
        Eliminar = True
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




Private Sub LanzaPantalla(Index As Integer)
Dim miI As Integer
        '----------------------------------------------------
        '----------------------------------------------------
        '
        ' Dependiendo de index lanzaremos una opcion uotra
        '
        '----------------------------------------------------
        
        'De momento solo para el 5. Cliente
        miI = -1
        Select Case Index
        Case 0
            txtaux(0).Text = ""
            miI = 3
        Case 3
            txtaux(3).Text = ""
            miI = 0
        Case 4
            txtaux(4).Text = ""
            miI = 1
            
        Case 8
            txtaux(8).Text = ""
            miI = 2
        End Select
       ' If miI >= 0 Then Image1_Click miI
End Sub





Private Function InsertarRegistro() As Boolean
On Error GoTo EInsertarLinea
    InsertarRegistro = False
    
    
    SQL = DevuelveDesdeBD("max(codigo)", "scarecepdoc", "1", "1") 'Truco del almendruco par obtener el max
    If SQL = "" Then SQL = "0"
    NumRegElim = Val(SQL) + 1
    
    
    Text1(4).Text = NumRegElim
    InsertarRegistro = InsertarDesdeForm(Me)
    
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, Err.Description
End Function



'A partir de un string empipado separaremos
Private Sub PonerDatosVencimiento(CADENA As String, Todo As Boolean)
    If RecuperaValor(CADENA, 6) = "0" Then
        If MsgBox("No tiene forma de pago correcta. �Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    If Todo Then
        txtaux(0).Text = RecuperaValor(CADENA, 1)
        txtaux(1).Text = RecuperaValor(CADENA, 2)
    End If
    txtaux(2).Text = Format(RecuperaValor(CADENA, 4), "dd/mm/yyyy")
    txtaux(3).Text = RecuperaValor(CADENA, 3)
    CADENA = RecuperaValor(CADENA, 5)
    ImporteVto = CCur(CADENA)
    txtaux(4).Text = Format(CADENA, FormatoImporte)
    
    Ponerfoco txtaux(4)
End Sub

'ImporteCoincide
'       0:  IMporte del tal/pag igual que el de la suma de las lineas
'       1:  Importe del  "       MAYOR  "
'       -1: Importe    "         MENOR  "
Private Sub HacerContabilizacion(ImporteCoincide As Integer)

    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(1).Text), True) > 1 Then Exit Sub
    
    'Cuenta bloqueada
    If CuentaBloqeada(Text1(2).Text, CDate(Text1(1).Text), True) Then Exit Sub
    
        
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
      
    
    'Abrireremos una ventana para seleccionar un par de cosillas
    If Combo1.ListIndex = 0 Then
        CadenaDesdeOtroForm = CStr(vbPagare)
    Else
        CadenaDesdeOtroForm = CStr(vbTalon)
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|" & CStr(ImporteCoincide) & "|"
    
    
    
    
    
    frmListado.Opcion = 23
    frmListado.Show vbModal



    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        CadAncho = RemesasCancelacionTALONPAGARE_(Combo1.ListIndex = 1, CInt(Text1(4).Text), CDate(Text1(1).Text), CadenaDesdeOtroForm)
        If CadAncho Then
            Conn.CommitTrans
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmActualizar2.OpcionActualizar = 20
            frmActualizar2.Show vbModal


            'Espera
            espera 0.2
            If SituarData1(True) Then PonerCampos
            

        Else
            TirarAtrasTransaccion
        End If
        CadAncho = True  'la vuelvo a poner como estaba
    End If
    Screen.MousePointer = vbDefault
End Sub


'Vamos a borrar el apunte generado anteriormente
'ImporteCoincide
'       0:  IMporte del tal/pag igual que el de la suma de las lineas
'       1:  Importe del  "       MAYOR  "
'       -1: Importe    "         MENOR  "
Private Function HacerDES_Contabilizacion_(ImporteCoincide As Integer) As Boolean

    
    HacerDES_Contabilizacion_ = False
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(1).Text), True) > 1 Then Exit Function
    
    'Cuenta bloqueada
    If CuentaBloqeada(Text1(2).Text, CDate(Text1(1).Text), True) Then Exit Function
    
        
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
      
    
    'Abrireremos una ventana para seleccionar un par de cosillas
    If Combo1.ListIndex = 0 Then
        CadenaDesdeOtroForm = CStr(vbPagare)
    Else
        CadenaDesdeOtroForm = CStr(vbTalon)
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|" & CStr(ImporteCoincide) & "|"
    
    
    
    
    
    frmListado.Opcion = 34
    frmListado.Show vbModal



    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        CadAncho = EliminarCancelacionTALONPAGARE(Combo1.ListIndex = 1, CInt(Text1(4).Text), CDate(Text1(1).Text), CadenaDesdeOtroForm)
        If CadAncho Then
            Conn.CommitTrans
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmActualizar2.OpcionActualizar = 20
            frmActualizar2.Show vbModal


            HacerDES_Contabilizacion_ = True
            
        Else
            TirarAtrasTransaccion
        End If
        CadAncho = True  'la vuelvo a poner como estaba
    End If
    Screen.MousePointer = vbDefault
End Function






Private Function MontaSQLDelVto(EnLasLineas As Boolean) As String
    If EnLasLineas Then
        MontaSQLDelVto = " numserie = '" & txtaux(0).Text & "' AND codfaccl = " & txtaux(1).Text
        MontaSQLDelVto = MontaSQLDelVto & " and fecfaccl ='" & Format(txtaux(2).Text, FormatoFecha) & "' AND numorden = " & txtaux(3).Text
    Else
        With Adodc1.Recordset
          MontaSQLDelVto = " numserie = '" & !NUmSerie & "' AND codfaccl = " & !numfaccl
          MontaSQLDelVto = MontaSQLDelVto & " and fecfaccl ='" & Format(!fecfaccl, FormatoFecha) & "' AND numorden = " & !numvenci
        End With
    End If
    
End Function

Private Function PuedeRealizarAccion(PermisoAdministrador As Boolean, ModificarCab As Boolean, Eliminar As Boolean) As Boolean
Dim TieneCtaPte As Boolean

    PuedeRealizarAccion = False
    If Data1.Recordset.EOF Then Exit Function
    If Modo <> 2 Then Exit Function
    If PermisoAdministrador Then
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permisos", vbExclamation
            Exit Function
        End If
    End If
    
    
    'AHora compruebo que no esta contabilizado
    SQL = DevuelveDesdeBD("LlevadoBanco", "scarecepdoc", "codigo", Text1(4).Text)
    If SQL = "1" Then
        'ESTA LLEVADA A BANCO
        If Combo1.ListIndex = 1 Then
            TieneCtaPte = vParam.TalonesCtaPuente
        Else
            TieneCtaPte = vParam.PagaresCtaPuente
        End If
        If Check1.Value = 0 And TieneCtaPte Then
            'Hay un error y no esta marcada como contabilziada
            MsgBox "Falta actualizar datos", vbExclamation
            PonerModo 0
            Exit Function
        End If
        
        
        SQL = DevuelveDesdeBD("Contabilizada", "scarecepdoc", "codigo", Text1(4).Text)
        If SQL = "0" Then
            If Not ModificarCab And TieneCtaPte Then
                MsgBox "Esta contabilizada pero no ha sido llevada a banco", vbExclamation
                Exit Function
            End If
        End If
        
        If ModificarCab Then
        
        Else
            If Not Eliminar Then
                MsgBox "Ya esta en banco", vbExclamation
                Exit Function
            End If
        
            'Si es eliminar
            SQL = "Select scobro.numserie,scobro.codfaccl,scobro.fecfaccl,scobro.numorden"
            SQL = SQL & " FROM slirecepdoc left join scobro on scobro.numserie=slirecepdoc.numserie AND codfaccl=numfaccl and"
            SQL = SQL & " scobro.fecfaccl = slirecepdoc.fecfaccl And numorden = numvenci"
            SQL = SQL & " WHERE id =" & Data1.Recordset!Codigo
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            SQL = ""
            NumRegElim = 0
            While Not miRsAux.EOF
                If Not IsNull(miRsAux!codfaccl) Then
                    SQL = SQL & DBLet(miRsAux!NUmSerie, "T") & Format(miRsAux!codfaccl, "000000") & "  " & Format(miRsAux!fecfaccl, "dd/mm/yyyy") & vbCrLf
                    NumRegElim = NumRegElim + 1
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            If NumRegElim > 0 Then
                'Hay vencimientos sin eliminar. No se pude eliminar el regisro
                If NumRegElim = 1 Then
                    SQL = "Existe un vencimiento pendiente de eliminar: " & vbCrLf & SQL
                Else
                    SQL = "Existen vencimientos(" & NumRegElim & ") pendientes de eliminar: " & vbCrLf & SQL
                End If
                MsgBox SQL, vbExclamation
                Exit Function
            End If
        End If


    Else
        'Si no esta llevada a banco
        'Si no es para modificar la cabecera si esta contabilizada TAMPOCO dejo continuar
        If Not ModificarCab Then
            'Para eliminar si que dejare pasar
            If Not Eliminar Then
                SQL = DevuelveDesdeBD("Contabilizada", "scarecepdoc", "codigo", Text1(4).Text)
                If SQL = "1" Then
                    'ESTA CONTABILIZADO
                    MsgBox "Esta contabilizada", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    PuedeRealizarAccion = True
    
End Function


Private Sub NuevoTalonPagareDefecto(Leer As Boolean)
Dim I As Integer
    On Error GoTo ENuevoTalonPagareDefecto
    If Leer Then
        I = CheckValueLeer("talpag")
        Me.Combo1.ListIndex = I
        
    Else
        'Escribir
        I = Combo1.ListIndex
        CheckValueGuardar "talpag", CByte(I)
    End If
    Exit Sub
ENuevoTalonPagareDefecto:
    Err.Clear

End Sub


Private Sub CambiaFechaVto()

    If Me.Data1.Recordset!fechavto <> CDate(Text1(6).Text) Then
        Set miRsAux = New ADODB.Recordset
        SQL = "SELECT numserie,numfaccl,fecfaccl,numvenci,importe FROm slirecepdoc WHERE id = " & Data1.Recordset!Codigo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = "UPDATE scobro set fecultco='" & Format(Text1(6).Text, FormatoFecha) & "' WHERE"
            SQL = SQL & " numserie = '" & miRsAux!NUmSerie & "' AND fecfaccl='" & Format(miRsAux!fecfaccl, FormatoFecha)
            SQL = SQL & "' AND numorden= " & miRsAux!numvenci & " AND codfaccl = " & miRsAux!numfaccl
            Ejecuta SQL
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub


Private Sub HacerBusqueda()

Dim CadB As String
CadB = ObtenerBusqueda(Me)

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

'Cuand esta a�adiendo una nueva, veremos si coinciden los importes
Private Function ComprobarImportes() As Boolean

On Error GoTo eComprobarImportes
    'Si ha ha sido llevada NO deberia haber entrado
    ComprobarImportes = True 'dejare que salga de las lineas
    SQL = DevuelveDesdeBD("LlevadoBanco", "scarecepdoc", "codigo", Text1(4).Text)
    
    If SQL = "1" Then
        MsgBox "No deberia haber entrado en edicion de lineas. Llevado a banco", vbExclamation
        Exit Function
    End If
    
    
    
    'Sumas lineas
    ImporteVto = 0
    SQL = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", Text1(4).Text)
    If SQL <> "" Then ImporteVto = CCur(SQL)
    SQL = Format(ImporteVto, FormatoImporte)
    
    
    If Me.Text1(5).Text <> SQL Then
        SQL = "Importes distintos: " & vbCrLf & "Talon/Pagar�: " & Text1(5).Text & vbCrLf & "Lineas vtos: " & SQL
        SQL = SQL & vbCrLf & vbCrLf & "�Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then ComprobarImportes = False
    End If
    
    Exit Function
eComprobarImportes:
    MuestraError Err.Number, Err.Description
End Function
