VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmColCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmColCtas2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   6450
      TabIndex        =   24
      Top             =   5100
      Width           =   1395
      Begin VB.OptionButton Option2 
         Caption         =   "Cod. Cta"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   600
         Width           =   1035
      End
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   6240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAccion 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   20
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   19
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   2220
      TabIndex        =   18
      Top             =   5040
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   1200
      TabIndex        =   17
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4515
      Left            =   6450
      TabIndex        =   10
      Top             =   600
      Width           =   1425
      Begin VB.CheckBox Check1 
         Caption         =   "9º nivel"
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   4245
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8º nivel"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   3805
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7º nivel"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   0
         Top             =   3365
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6º nivel"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2925
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5º nivel"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2485
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4º nivel"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   2045
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3º nivel"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1620
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2º nivel"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1165
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1er nivel"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   725
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Último:  "
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   285
         Value           =   1  'Checked
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColCtas2.frx":000C
      Height          =   5415
      Left            =   90
      TabIndex        =   12
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6030
      Top             =   30
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
      TabIndex        =   15
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Busqueda avanzada"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Comprobar cuentas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5160
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label lblComprobar 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   22
      Top             =   6300
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblComprobar 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   21
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Menu mnOpciones 
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmColCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Busqueda As String   'Para que ponga la cadena donde le corresponde
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public ConfigurarBalances As Byte
    '0.- Normal
    '1.- Busqueda
    '2.- Agrupacion de cuentas
    '3.- BUSQUEDA NUEVA
    '4.- Nueva cuenta
    '5.- Busquedas de envio de e-mail
    '6.- Exclusion de cuentas en consolidado. Como la agrupacion pero acepta niveles inferiores al penultimo
    
    
    '7.- Lleva una cadnea de texto para buscar
    
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Dim CadAncho As Boolean 'Para cuando llamemos al al form de lineas
Dim RS As Recordset
Dim NF As Integer
Dim Errores As Long
Dim PrimeraVez As Boolean

Private Sub BotonAnyadir(Cuenta As String)
'    ParaBusqueda False
'    frmCuentas.vModo = 1
'    frmCuentas.CodCta = Cuenta
'    CadenaDesdeOtroForm = ""
'    frmCuentas.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        Screen.MousePointer = vbHourglass
'        CargaGrid
'        'Intentamos situar el grid
'        SituaGrid CadenaDesdeOtroForm
'        If Me.ConfigurarBalances = 4 Then cmdRegresar_Click
'    Else
'        If Me.ConfigurarBalances = 4 Then CargaGrid
'    End If
End Sub

Private Sub BotonBuscar()
    CadenaConsulta = GeneraSQL("codmacta= 'David'")  'esto es para que no cargue ningun registro
    CargaGrid
    ParaBusqueda True
    txtaux(0).Text = "": txtaux(1).Text = ""
    Ponerfoco txtaux(1)
End Sub

Private Sub BotonVerTodos()
    ParaBusqueda False
    CadenaConsulta = GeneraSQL("")
    CargaGrid
End Sub



Private Sub BotonModificar()
'    ParaBusqueda False
'    CadenaDesdeOtroForm = ""
'    frmCuentas.vModo = 2
'    frmCuentas.CodCta = Adodc1.Recordset!codmacta
'    frmCuentas.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        CargaGrid
'        SituaGrid CadenaDesdeOtroForm
'    End If
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    ParaBusqueda False
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar la cuenta:"
    SQL = SQL & vbCrLf & "Código: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = Adodc1.Recordset.Fields(0)
        If SepuedeEliminarCuenta(SQL) Then
            'Hay que eliminar
            Screen.MousePointer = vbHourglass
            SQL = "Delete from cuentas where codmacta='" & Adodc1.Recordset!codmacta & "'"
            Conn.Execute SQL
            Screen.MousePointer = vbHourglass
            espera 0.5
            'Cancelamos el adodc1
            DataGrid1.Enabled = False
            CargaGrid
            DataGrid1.Enabled = True
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar cuenta contable"
End Sub

Private Sub Check1_Click(Index As Integer)
    OpcionesCambiadas
End Sub

Private Sub OpcionesCambiadas()
    Screen.MousePointer = vbHourglass
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim SQL As String
Dim Aux As String
    If Index = 0 Then
        'Ha pulsado aceptar
        txtaux(0).Text = Trim(txtaux(0).Text)
        txtaux(1).Text = Trim(txtaux(1).Text)
        'Si estan vacios no hacemos nada
        SQL = ""
        Aux = ""
        If txtaux(0).Text <> "" Then
            If SeparaCampoBusqueda("T", "codmacta", txtaux(0).Text, Aux) = 0 Then SQL = Aux
        End If
        If txtaux(1).Text <> "" Then
            Aux = ""
            If InStr(1, txtaux(1).Text, "*") = 0 Then txtaux(1).Text = "*" & txtaux(1).Text & "*"
            If SeparaCampoBusqueda("T", "nommacta", txtaux(1).Text, Aux) = 0 Then
                If SQL <> "" Then SQL = SQL & " AND "
                SQL = SQL & Aux
            End If
        End If
        
        'Si sql<>"" entonces hay puestos valores
        If SQL = "" Then Exit Sub
        
        'Llamamos a carga grid
        Screen.MousePointer = vbHourglass
        CadenaConsulta = GeneraSQL(SQL)
        CargaGrid
        Screen.MousePointer = vbDefault
        If Adodc1.Recordset.EOF Then
            MsgBox "Ningún resultado para la búsqueda.", vbExclamation
            Exit Sub
        End If
        Ponerfoco DataGrid1
    End If
    ParaBusqueda False
    lblIndicador.Caption = ""
End Sub

Private Sub cmdRegresar_Click()
    
    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    RaiseEvent DatoSeleccionado(Adodc1.Recordset!codmacta & "|" & Adodc1.Recordset!Nommacta & "|")
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then
        cmdRegresar_Click
    Else
        'Vemos todos los valores de la cuenta
'        frmCuentas.vModo = 0
'        frmCuentas.CodCta = Adodc1.Recordset!codmacta
'        frmCuentas.Show vbModal
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerOptionsVisibles 1
        
        'Vamos a ver si funciona 30 Sept 2003
        Select Case ConfigurarBalances
        Case 0, 1, 2, 5
            If ConfigurarBalances = 5 Then
                'Estoy buscando los que tienen e-mail
                CadenaConsulta = CadenaConsulta & " AND maidatos <> ''"
            End If
            CargaGrid
        Case 3
            BotonBuscar
        Case 4
            
            BotonAnyadir CadenaDesdeOtroForm
        Case 7
            CadenaConsulta = CadenaConsulta & " AND codmacta like '" & Busqueda & "%'"
            CargaGrid
        End Select
        CadenaDesdeOtroForm = ""
    End If
    Screen.MousePointer = vbDefault
End Sub


'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
     
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 24
        .Buttons(3).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 20
        .Buttons(14).Image = 15
    End With
     
    PrimeraVez = True
    pb1.Visible = False
    'Poner niveles
    PonerOptionsVisibles 0
    
    'Opciones segun sea su nivel
    PonerOpcionesMenu
    
    'Ocultamos busqueda
    ParaBusqueda False
    
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.Password = vUsu.Passwd
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    Frame2.Enabled = (DatosADevolverBusqueda = "")
    CadAncho = False
    'Cadena consulta
    CadenaConsulta = GeneraSQL("")
    
    lblIndicador.Caption = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    ConfigurarBalances = 0
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
    BotonAnyadir ""
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub



'----------------------------------------------------------------


Private Sub Option2_Click(Index As Integer)
    OpcionesCambiadas
End Sub

Private Sub Option2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
            BotonBuscar
    Case 2
        'Busqueda avanzada
'            CadenaDesdeOtroForm = ""
'            frmCuentas.vModo = 3
'            frmCuentas.Show vbModal
'            If CadenaDesdeOtroForm <> "" Then
'                Me.Refresh
'                Screen.MousePointer = vbHourglass
'                PonerResultadosBusquedaAvanzada
'                Screen.MousePointer = vbDefault
'            End If
    Case 3
            BotonVerTodos
    Case 6
            BotonAnyadir ""
    Case 7
            BotonModificar
    Case 8
            BotonEliminar
    Case 11
            'Imprimimos el listado
'                Screen.MousePointer = vbHourglass
'                frmListado.Opcion = 2 'Listado de cuentas
'                frmListado.Show vbModal
    
    Case 12
            'Comprobar cuentas
            ComprobarCuentas
    Case 14
            Unload Me
    Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 16 To 19
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub


Private Sub CargaGrid()
Dim B As Boolean
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    CargaGrid2
    DataGrid1.Enabled = B
End Sub

Private Sub CargaGrid2()
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    Dim SQL As String
    Dim B As Boolean
    Adodc1.ConnectionString = Conn
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    SQL = CadenaConsulta
    SQL = SQL & " ORDER BY"
    If Option2(0).Value Then
        SQL = SQL & " codmacta"
    Else
        SQL = SQL & " nommacta"
    End If
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    
   
        DataGrid1.Columns(0).Caption = "Cuenta"
        DataGrid1.Columns(0).Width = 1200
    
   
        DataGrid1.Columns(1).Caption = "Denominación"
        DataGrid1.Columns(1).Width = 4000
        TotalAncho = TotalAncho + DataGrid1.Columns(1).Width
    
   
        DataGrid1.Columns(2).Caption = "Direc."
        DataGrid1.Columns(2).Width = 500
        TotalAncho = TotalAncho + DataGrid1.Columns(2).Width
               
        If Not CadAncho Then
            txtaux(0).Left = DataGrid1.Columns(0).Left + 100
            txtaux(0).Width = DataGrid1.Columns(0).Width - 15
            txtaux(0).Top = DataGrid1.Top + 235
            txtaux(1).Left = DataGrid1.Columns(1).Left + 100
            txtaux(1).Width = DataGrid1.Columns(1).Width - 15
            txtaux(1).Top = txtaux(0).Top
            txtaux(0).Height = DataGrid1.RowHeight - 15
            txtaux(1).Height = txtaux(0).Height
            CadAncho = True
        End If
               
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    End If
   
    'Para k la barra de desplazamiento sea mas alta
    If Not Adodc1.Recordset.EOF Then
            DataGrid1.ScrollBars = dbgVertical
    End If
    DataGrid1.Enabled = B
End Sub


' 0 solo textos
'1 Solo enables
'2 todo
Private Sub PonerOptionsVisibles(Opcion As Byte)
Dim I As Integer
Dim J As Integer
Dim Cad As String

    'Utilizo la variable cadancho
If Opcion <> 1 Then
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        If J > 0 Then
            Cad = "Digitos: " & J
            Check1(I).Caption = Cad
        Else
            Check1(I).Caption = "Error"
        End If
        Check1(I).Value = 0
    Next I
    'Ultimo nivel
    J = DigitosNivel(I)
    If J > 0 Then Check1(0).Caption = Check1(0).Caption & J
    For I = vEmpresa.numnivel To 9
        Check1(I).Visible = False
    Next I
End If
If Opcion <> 0 Then
    Select Case ConfigurarBalances
    Case 1
        For I = 1 To vEmpresa.numnivel - 1
            J = DigitosNivel(I)
            If J < 5 Then  'A balances van ctas de 4 digitos
               Check1(I).Value = 1
            Else
               Check1(I).Value = 0
            End If
        Next I
        Check1(0).Value = 1
    Case 2
        'Agrupar ctas digitos . Realmete agrupamos al nivel de cuentas -1
'        J = DevuelveDigitosNivelAnterior
'        For I = 0 To 9
'            Check1(I).Visible = I = J
'            Check1(I).Value = Abs(I = J)
'        Next I
        
    Case 6
        'Todos los niveles menos el ultimo
        'Agrupar ctas digitos . Realmete agrupamos al nivel de cuentas -1
'        J = DevuelveDigitosNivelAnterior
'        For I = 0 To 9
'            Check1(I).Visible = I < J And I > 0
'            Check1(I).Value = Abs(I <= J) And I > 0
'        Next I
    Case Else
        Check1(0).Value = 1
    End Select
        
End If
End Sub



Private Function GeneraSQL(Busqueda As String) As String
Dim I As Integer
Dim SQL As String
Dim nexo As String
Dim J As Integer
Dim wildcar As String

SQL = ""
nexo = ""
If Check1(0).Value Then
    SQL = "( apudirec = 'S')"
    nexo = " OR "
End If
For I = 1 To vEmpresa.numnivel - 1
    If Check1(I).Value = 1 Then
        wildcar = ""
        For J = 1 To I
            wildcar = wildcar & "_"
        Next J
        SQL = SQL & nexo & " ( codmacta like '" & wildcar & "')"
        nexo = " OR "
    End If
Next I
wildcar = "SELECT codmacta, nommacta, apudirec"
wildcar = wildcar & " FROM cuentas "


'Nexo
nexo = " WHERE "
If Busqueda <> "" Then
    wildcar = wildcar & " WHERE (" & Busqueda & ")"
    nexo = " AND "
End If
If SQL <> "" Then wildcar = wildcar & nexo & "(" & SQL & ")"

GeneraSQL = wildcar
End Function



Private Function SepuedeEliminarCuenta(Cuenta As String) As Boolean
Dim NivelCta As Integer
Dim I, J As Integer
Dim Cad As String

    SepuedeEliminarCuenta = False
    If EsCuentaUltimoNivel(Cuenta) Then
        'ATENCION###
        ' Habra que ver casos particulares de eliminacion de una subcuenta de ultimo nivel
        'Si esta en apuntes, en ....
        'NO se puede borrar
        If Not BorrarCuenta(Cuenta) Then Exit Function
    Else
        'No
        'No
        'no es una cuenta de ultimo nivel
        NivelCta = NivelCuenta(Cuenta)
        If NivelCta < 1 Then
            MsgBox "Error obteniendo nivel de la subcuenta", vbExclamation
            Exit Function
        End If
        
        'Ctas agrupadas
        I = DigitosNivel(NivelCta)
        If I = 3 Then
            Cad = DevuelveDesdeBD("codmacta", "ctaagrupadas", "codmacta", Cuenta, "T")
            If Cad <> "" Then
                MsgBox "El subnivel pertenece a agrupacion de cuentas en balance"
                Exit Function
            End If
        End If
        For J = NivelCta + 1 To vEmpresa.numnivel
            Cad = Cuenta & "__________"
            I = DigitosNivel(J)
            Cad = Mid(Cad, 1, I)
            If TieneEnBD(Cad) Then
                MsgBox "Tiene cuentas en niveles superiores (" & J & ")", vbExclamation
                Exit Function
            End If
        Next J
    End If
    SepuedeEliminarCuenta = True
End Function

Private Function TieneEnBD(Cad As String) As String
    'Dim Cad1 As String
    
    Set RS = New ADODB.Recordset
    'Cad1 = "Select codmacta from cuentas where codmacta like '" & Cad & "'"
    RS.Open "Select codmacta from cuentas where codmacta like '" & Cad & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TieneEnBD = Not RS.EOF
    RS.Close
    Set RS = Nothing
End Function


Private Function BorrarCuenta(Cuenta As String) As Boolean
On Error GoTo Salida
Dim SQL As String


pb1.Max = 6
pb1.Value = 0
pb1.Visible = True

'Con ls tablas declarads sin el ON DELETE , no dejara borrar
BorrarCuenta = False
Set RS = New ADODB.Recordset


'lineas de apuntes, contrapartidads   -->1
RS.Open "Select * from linasipre where ctacontr ='" & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    RS.Close
    GoTo Salida
End If
RS.Close
pb1.Value = 1

'-->2
RS.Open "Select * from linapu where ctacontr ='" & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    RS.Close
    GoTo Salida
End If
RS.Close
pb1.Value = 2

'-->3
'Otras tablas
'Reparto de gastos para inmovilizado
SQL = "Select codmacta2 from sbasin where codmacta2='" & Cuenta & "'"
RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    RS.Close
    GoTo Salida
End If
RS.Close
pb1.Value = 3



'-->4
RS.Open "Select * from presupuestos where codmacta ='" & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    RS.Close
    GoTo Salida
End If
RS.Close
pb1.Value = 4


'-->5    Referencias a ctas desde eltos de inmovilizado
SQL = "select codinmov from sinmov where codmact1='" & Cuenta & "'"
SQL = SQL & " or codmact2='" & Cuenta & "'"
SQL = SQL & " or codmact3='" & Cuenta & "'"
RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    RS.Close
    GoTo Salida
End If
RS.Close
pb1.Value = 5



'-->6    Referencias a ctas desde eltos de inmovilizado
SQL = "select codiva from samort where codiva='" & Cuenta & "'"
RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    RS.Close
    GoTo Salida
End If
RS.Close
pb1.Value = 6


'SI kkega aqui es k ha ido bien
BorrarCuenta = True
Salida:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar ctas." & Err.Description
    Set RS = Nothing
    pb1.Visible = False
End Function


Private Sub SituaGrid(CADENA As String)
On Error GoTo ESituaGrid
If Adodc1.Recordset.EOF Then Exit Sub

Adodc1.Recordset.Find " codmacta =  " & CADENA & ""
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst

Exit Sub
ESituaGrid:
    MuestraError Err.Number, "Situando registro activo"
End Sub


Private Sub ParaBusqueda(Ver As Boolean)
txtaux(0).Visible = Ver
txtaux(1).Visible = Ver
cmdAccion(0).Visible = Ver
cmdAccion(1).Visible = Ver
If Ver Then lblIndicador.Caption = "Búsqueda"
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtaux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub ComprobarCuentas()
Dim Cad As String
Dim N As Integer
Dim I As Integer

On Error GoTo EComprobarCuentas
'NO hay cuentas
If Me.Adodc1.Recordset.EOF Then Exit Sub
'Buscando datos
If txtaux(0).Visible Then Exit Sub
'Para cada nivel n comprobaremos si existe la cuenta en un
'nivel n-1
'La comprobacion se hara para cada cta de n sabiendo k
' para las cuentas de nivel 4 digitos  4300 ..4309 tienen
'el mismo subnivel n-1 430
lblComprobar(0).Caption = ""
lblComprobar(1).Caption = ""
lblComprobar(0).Visible = True
lblComprobar(1).Visible = True
Me.lblIndicador.Caption = "Comprobar cuentas"
Me.Refresh
Errores = 0
NF = FreeFile
Open App.Path & "\Errorcta.txt" For Output As #NF


'Primero comprobamos las cuentas de mayor longitud que la permitida
CuentasDeMasNivel
lblComprobar(0).Caption = "Cuentas > Ultimo nivel"
lblComprobar(0).Refresh
'Hasta 2 pq el uno no tiene subniveles
For I = vEmpresa.numnivel To 2 Step -1
    N = DigitosNivel(I)
    lblComprobar(0).Caption = "Nivel: " & I
    lblComprobar(0).Refresh
    Do
        If ObtenerCuenta(Cad, I, N) Then
            lblComprobar(1).Caption = Cad
            lblComprobar(1).Refresh
            ComprobarCuenta Cad, I
        End If
    Loop Until Cad = ""
Next I

Close #NF
If Errores = 0 Then
    Kill App.Path & "\Errorcta.txt"
    Else
        Cad = Dir("C:\WINDOWS\NOTEPAD.exe")
        If Cad = "" Then
            Cad = Dir("C:\WINNT\NOTEPAD.exe")
        End If
        If Cad = "" Then
            MsgBox "Se ha producido errores. Vea el archivo Errorcta.txt"
            Else
            Shell Cad & " " & App.Path & "\Errorcta.txt", vbMaximizedFocus
        End If
End If

EComprobarCuentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar cuentas: ", Err.Description
        Close #NF
    End If
    Me.lblComprobar(0).Visible = False
    Me.lblComprobar(1).Visible = False
    Me.lblIndicador.Caption = ""
    Me.Refresh
End Sub


Private Function ObtenerCuenta(ByRef CADENA As String, Nivel As Integer, ByRef Digitos As Integer) As Boolean
Dim RT As Recordset
Dim SQL As String


If CADENA = "" Then
    SQL = ""
Else
    SQL = DevuelveUltimaCuentaGrupo(CADENA, Nivel, Digitos)
    SQL = " codmacta > '" & SQL & "' AND "
End If
SQL = "Select codmacta from Cuentas WHERE " & SQL
SQL = SQL & " codmacta like '" & Mid("__________", 1, Digitos) & "'"

Set RT = New ADODB.Recordset
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RT.EOF Then
    ObtenerCuenta = False
    CADENA = ""
Else
    ObtenerCuenta = True
    CADENA = RT!codmacta
End If
RT.Close
Set RT = Nothing
End Function


Private Sub ComprobarCuenta(Cuenta As String, Nivel As Integer)
Dim N As Integer
Dim Aux As String
Dim aux2 As String

N = DigitosNivel(Nivel - 1)
Aux = Mid(Cuenta, 1, N)
aux2 = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Aux, "T")
If aux2 = "" Then
    'Error
   Errores = Errores + 1
   Print #NF, "Nivel: " & Nivel
   Print #NF, "Cuenta: " & Cuenta & "  -> " & Aux & " NO encontrada "
   Print #NF, ""
   Print #NF, ""
End If

End Sub



Private Function DevuelveUltimaCuentaGrupo(Cta As String, Nivel As Integer, ByRef Digitos As Integer) As String
Dim Cad As String
Dim N As Integer
N = DigitosNivel(Nivel - 1)
Cad = Mid(Cta, 1, N)
Cad = Cad & "9999999999"
DevuelveUltimaCuentaGrupo = Mid(Cad, 1, Digitos)
End Function


Private Sub CuentasDeMasNivel()
'###MYSQL
Set RS = New ADODB.Recordset
RS.Open "SELECT codmacta FROM cuentas WHERE ((Length(cuentas.codmacta)>" & vEmpresa.DigitosUltimoNivel & "))", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    Print #NF, "Cuentas de longitud mayor a la permitida"
    Print #NF, "Digitos ultimo nivel: " & vEmpresa.DigitosUltimoNivel
    While Not RS.EOF
        Errores = Errores + 1
        Print #NF, "      .- " & RS!codmacta
        RS.MoveNext
    Wend
    Print #NF, ""
    Print #NF, ""
End If
RS.Close
Set RS = Nothing
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    'Caption = KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then Unload Me
    End If
End Sub


Private Sub Ponerfoco(ByRef Obje As Object)
On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerResultadosBusquedaAvanzada()

    On Error GoTo EC
        CadenaConsulta = GeneraSQL(CadenaDesdeOtroForm)
        CargaGrid
    Exit Sub
EC:
    MuestraError Err.Number, "Poner resultados busqueda avanzada"
End Sub
