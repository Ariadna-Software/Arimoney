VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFormaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de pago"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frmfromaPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmfromaPago.frx":030A
      Left            =   5640
      List            =   "frmfromaPago.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Tipo forma pago|N|N|0||sforpa|tipforpa|||"
      Top             =   720
      Width           =   3180
   End
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1380
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||sforpa|nomforpa|||"
      Text            =   "Text1"
      Top             =   750
      Width           =   3285
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "Codigo|N|N|0||sforpa|codforpa||S|"
      Text            =   "Text1"
      Top             =   750
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   6000
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   390
      Top             =   6240
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
      TabIndex        =   11
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
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
         Left            =   5280
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo forma de pago"
      Height          =   195
      Left            =   5640
      TabIndex        =   10
      Top             =   480
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Denominacion"
      Height          =   255
      Index           =   1
      Left            =   1365
      TabIndex        =   9
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
'
'
'Private WithEvents frmB As frmBuscaGrid
''-----------------------------
''Se distinguen varios modos
''   0.-  Formulario limpio sin nungun campo rellenado
''   1.-  Preparando para hacer la busqueda
''   2.-  Ya tenemos registros y los vamos a recorrer
''        y podemos editarlos Edicion del campo
''   3.-  Insercion de nuevo registro
''   4.-  Modificar
''-------------------------------------------------------------------------
''-------------------------------------------------------------------------
''  Variables comunes a todos los formularios
'Private Modo As Byte
'Private CadenaConsulta As String
'Private Ordenacion As String
'Private NombreTabla As String  'Nombre de la tabla o de la
'Private kCampo As Integer
''-------------------------------------------------------------------------
'Private HaDevueltoDatos As Boolean
'Private DevfrmCCtas As String
'
'
'
'Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub cmdAceptar_Click()
'    Dim Cad As String
'    Dim I As Integer
'
'    Screen.MousePointer = vbHourglass
'    On Error GoTo Error1
'    Select Case Modo
'    Case 3
'        If DatosOk Then
'            '-----------------------------------------
'            'Hacemos insertar
'            If InsertarDesdeForm(Me) Then
'                'MsgBox "Registro insertado.", vbInformation
'                PonerModo 0
'                lblIndicador.Caption = ""
'            End If
'        End If
'    Case 4
'            'Modificar
'            If DatosOk Then
'                '-----------------------------------------
'                'Hacemos insertar
'                If ModificaDesdeFormulario(Me) Then
'                    TerminaBloquear
'                    lblIndicador.Caption = ""
'                    If SituarData1 Then
'                        PonerModo 2
'                    Else
'                        LimpiarCampos
'                        PonerModo 0
'                    End If
'                End If
'            End If
'    Case 1
'        HacerBusqueda
'    End Select
'
'Error1:
'        Screen.MousePointer = vbDefault
'        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
'End Sub
'
'Private Sub cmdCancelar_Click()
'Select Case Modo
'Case 1, 3
'    LimpiarCampos
'    PonerModo 0
'Case 4
'    'Modificar
'    lblIndicador.Caption = ""
'    TerminaBloquear
'    PonerModo 2
'    PonerCampos
'End Select
'
'End Sub
'
'
'' Cuando modificamos el data1 se mueve de lugar, luego volvemos
'' ponerlo en el sitio
'' Para ello con find y un SQL lo hacemos
'' Buscamos por el codigo, que estara en un text u  otro
'' Normalmente el text(0)
'Private Function SituarData1() As Boolean
'    Dim SQL As String
'    On Error GoTo ESituarData1
'            'Actualizamos el recordset
'            Data1.Refresh
'            '#### A mano.
'            'El sql para que se situe en el registro en especial es el siguiente
'            SQL = " codforpa = " & Text1(0).Text & ""
'            Data1.Recordset.Find SQL
'            If Data1.Recordset.EOF Then GoTo ESituarData1
'            SituarData1 = True
'        Exit Function
'ESituarData1:
'        If Err.Number <> 0 Then Err.Clear
'        Limpiar Me
'        PonerModo 0
'        lblIndicador.Caption = ""
'        SituarData1 = False
'End Function
'
'Private Sub BotonAnyadir()
'    LimpiarCampos
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    cmdAceptar.Caption = "Aceptar"
'    PonerModo 3
'    'Escondemos el navegador y ponemos insertando
'    DespalzamientoVisible False
'    lblIndicador.Caption = "INSERTANDO"
'    SugerirCodigoSiguiente
'    '###A mano
'    PonerFoco Text1(0)
'End Sub
'
'Private Sub BotonBuscar()
'    'Buscar
'    If Modo <> 1 Then
'        LimpiarCampos
'        lblIndicador.Caption = "Búsqueda"
'        PonerModo 1
'        '### A mano
'        '################################################
'        'Si pasamos el control aqui lo ponemos en amarillo
'        Text1(0).SetFocus
'        Text1(0).BackColor = vbYellow
'        Else
'            HacerBusqueda
'            If Data1.Recordset.EOF Then
'                 '### A mano
'                Text1(kCampo).Text = ""
'                Text1(kCampo).BackColor = vbYellow
'                Text1(kCampo).SetFocus
'            End If
'    End If
'End Sub
'
'Private Sub BotonVerTodos()
'    'Ver todos
'    LimpiarCampos
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
'        CadenaConsulta = "Select * from " & NombreTabla
'        PonerCadenaBusqueda
'    End If
'End Sub
'
'Private Sub Desplazamiento(Index As Integer)
'Select Case Index
'    Case 0
'        Data1.Recordset.MoveFirst
'    Case 1
'        Data1.Recordset.MovePrevious
'        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
'    Case 2
'        Data1.Recordset.MoveNext
'        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
'    Case 3
'        Data1.Recordset.MoveLast
'End Select
'PonerCampos
'lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
'End Sub
'
'Private Sub BotonModificar()
'    '---------
'    'MODIFICAR
'    '----------
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'   ' cmdAceptar.Caption = "Modificar"
'    PonerModo 4
'    'Escondemos el navegador y ponemos insertando
'    'Como el campo 1 es clave primaria, NO se puede modificar
'    '### A mano
'    Text1(0).Locked = True
'    Text1(0).BackColor = &H80000018
'    DespalzamientoVisible False
'    lblIndicador.Caption = "Modificar"
'End Sub
'
'Private Sub BotonEliminar()
'    Dim Cad As String
'    Dim I As Integer
'
'    'Ciertas comprobaciones
'    If Data1.Recordset.EOF Then Exit Sub
'
'    'Comprobamos si se puede eliminar
'    If Not SePuedeEliminar Then Exit Sub
'    '### a mano
'    Cad = "Seguro que desea eliminar de la BD el registro:"
'    I = MsgBox(Cad, vbQuestion + vbYesNo)
'    'Borramos
'    If I = vbYes Then
'        'Hay que eliminar
'        On Error GoTo Error2
'        Screen.MousePointer = vbHourglass
'        NumRegElim = Data1.Recordset.AbsolutePosition
'        Data1.Recordset.Delete
'        Data1.Refresh
'        If Data1.Recordset.EOF Then
'            'Solo habia un registro
'            LimpiarCampos
'            PonerModo 0
'            Else
'                Data1.Recordset.MoveFirst
'                NumRegElim = NumRegElim - 1
'                If NumRegElim > 1 Then
'                    For I = 1 To NumRegElim - 1
'                        Data1.Recordset.MoveNext
'                    Next I
'                End If
'                PonerCampos
'        End If
'    End If
'Error2:
'        Screen.MousePointer = vbDefault
'        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
'End Sub
'
'
'
'
'Private Sub cmdRegresar_Click()
'Dim Cad As String
'Dim I As Integer
'Dim J As Integer
'Dim Aux As String
'
'If Data1.Recordset.EOF Then
'    MsgBox "Ningún registro devuelto.", vbExclamation
'    Exit Sub
'End If
'
'Cad = ""
'I = 0
'Do
'    J = I + 1
'    I = InStr(J, DatosADevolverBusqueda, "|")
'    If I > 0 Then
'        Aux = Mid(DatosADevolverBusqueda, J, I - J)
'        J = Val(Aux)
'        Cad = Cad & Text1(J).Text & "|"
'    End If
'Loop Until I = 0
'RaiseEvent DatoSeleccionado(Cad)
'Unload Me
'End Sub
'
'Private Sub KEYpress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Combo1.BackColor = vbWhite
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'
'Private Sub Combo1_KeyPress(KeyAscii As Integer)
'        KEYpress KeyAscii
'End Sub
'
'
'Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub Form_Activate()
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub Form_Load()
'Dim I As Integer
'
'
'      ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1
'        .Buttons(2).Image = 2
'        .Buttons(6).Image = 3
'        .Buttons(7).Image = 4
'        .Buttons(8).Image = 5
'        '.Buttons(10).Image = 10
'        .Buttons(11).Image = 16
'        .Buttons(12).Image = 15
'        .Buttons(14).Image = 6
'        .Buttons(15).Image = 7
'        .Buttons(16).Image = 8
'        .Buttons(17).Image = 9
'    End With
'
'
'    LimpiarCampos
'    'Si hay algun combo los cargamos
'    CargarCombo
'
'
'    '## A mano
'    NombreTabla = "sforpa"
'    Ordenacion = " ORDER BY codforpa"
'
'    PonerOpcionesMenu
'
'    'Para todos
''    Data1.UserName = vUsu.Login
''    Me.Data1.password = vUsu.Passwd
'    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
'    'ASignamos un SQL al DATA1
'    Data1.ConnectionString = Conn
'    Data1.RecordSource = "Select * from " & NombreTabla
'    Data1.Refresh
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'        Else
'        PonerModo 1
'        '### A mano
'        Text1(0).BackColor = vbYellow
'    End If
'
'End Sub
'
'
'
'Private Sub LimpiarCampos()
'    Limpiar Me   'Metodo general
'    lblIndicador.Caption = ""
'    'Aqui va el especifico de cada form es
'    '### a mano
'    Combo1.ListIndex = -1
'
'End Sub
'
'
'Private Sub CargarCombo()
'Dim RS As ADODB.Recordset
''###
''Cargaremos el combo, o bien desde una tabla o con valores fijos o como
''se quiera, la cuestion es cargarlo
'' El estilo del combo debe de ser 2 - Dropdown List
'' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
'' o marcamos la opcion sorted del combo
'
'    Combo1.Clear
'
'    Set RS = New ADODB.Recordset
'    DevfrmCCtas = "SELECT * FROm stipoformapago ORDER BY tipoformapago"
'    RS.Open DevfrmCCtas, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RS.EOF
'        Combo1.AddItem RS!descformapago
'        Combo1.ItemData(Combo1.NewIndex) = RS!tipoformapago
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
'End Sub
'
'Private Sub frmB_Selecionado(CadenaDevuelta As String)
'    If CadenaDevuelta <> "" Then
'
'        If DevfrmCCtas <> "" Then
'
'            HaDevueltoDatos = True
'            DevfrmCCtas = CadenaDevuelta
'
'        Else
'                HaDevueltoDatos = True
'                Screen.MousePointer = vbHourglass
'                'Sabemos que campos son los que nos devuelve
'                'Creamos una cadena consulta y ponemos los datos
'                DevfrmCCtas = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
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
'        End If
'    Else
'        DevfrmCCtas = ""
'    End If
'End Sub
'
'Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
'DevfrmCCtas = CadenaSeleccion
'End Sub
'
'Private Sub imgCuentas_Click(Index As Integer)
'Dim Cad As String
'
'
' Screen.MousePointer = vbHourglass
'
'
'
'
' Select Case Index
' Case 2, 10
'    'Diario
'        DevfrmCCtas = "0"
'        Cad = "Número|numdiari|N|30·"
'        Cad = Cad & "Descripción|desdiari|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = "Tiposdiario"
'        frmB.vSQL = ""
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Diario"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'           Text1(Index) = RecuperaValor(DevfrmCCtas, 1)
'           Text2(Index) = RecuperaValor(DevfrmCCtas, 2)
'        End If
' Case 3, 4, 8, 11
'        'Conceptos
'        DevfrmCCtas = "0"
'        Cad = "Codigo|codconce|N|30·"
'        Cad = Cad & "Descripción|nomconce|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = "Conceptos"
'        frmB.vSQL = ""
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "CONCEPTOS"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'           Text1(Index) = RecuperaValor(DevfrmCCtas, 1)
'           Text2(Index) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
' Case Else
'
'
'
' End Select
'
'
'
'End Sub
'
'
'Private Sub mnBuscar_Click()
'    BotonBuscar
'End Sub
'
'Private Sub mnEliminar_Click()
'    BotonEliminar
'End Sub
'
'Private Sub mnModificar_Click()
'    BotonModificar
'End Sub
'
'Private Sub mnNuevo_Click()
'BotonAnyadir
'End Sub
'
'Private Sub mnSalir_Click()
'Screen.MousePointer = vbHourglass
'Unload Me
'End Sub
'
'Private Sub mnVerTodos_Click()
'BotonVerTodos
'End Sub
'
'
''### A mano
''Los metodos del text tendran que estar
''Los descomentamos cuando esten puestos ya los controles
'Private Sub Text1_GotFocus(Index As Integer)
'    kCampo = Index
'    If Modo = 1 Then
'        Text1(Index).BackColor = vbYellow
'        Else
'            Text1(Index).SelStart = 0
'            Text1(Index).SelLength = Len(Text1(Index).Text)
'    End If
'End Sub
'
'
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
'End Sub
'
''----------------------------------------------------------------
''----------------------------------------------------------------
'' Cunado el campo de texto pierde el enfoque
'' Es especifico de cada formulario y en el podremos controlar
'' lo que queramos, desde formatear un campo si asi lo deseamos
'' hasta pedir que nos devuelva los datos de la empresa
''----------------------------------------------------------------
''----------------------------------------------------------------
'Private Sub Text1_LostFocus(Index As Integer)
'    Dim SQL As String
'
'    ''Quitamos blancos por los lados
'    Text1(Index).Text = Trim(Text1(Index).Text)
'    If Text1(Index).BackColor = vbYellow Then
'        Text1(Index).BackColor = &H80000018
'    End If
'
'
'
'
'    'Si queremos hacer algo ..
'    Select Case Index
'        Case 2, 10
'
'            If Modo = 3 Or Modo = 4 Then
'                'Insertando
'
'                If Text1(Index).Text = "" Then
'                    Text2(Index).Text = ""
'                    Exit Sub
'                End If
'                SQL = ""
'                If Not IsNumeric(Text1(Index).Text) Then
'                    MsgBox "Tipo de diario no es numérico: " & Text1(Index).Text, vbExclamation
'                    Text1(Index).Text = ""
'                    Text2(Index).Text = ""
'                    PonerFoco Text1(Index)
'                    Exit Sub
'                Else
'                    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(Index).Text, "N")
'                End If
'                If SQL = "" Then
'                    SQL = "Diario no encontrado: " & Text1(Index).Text
'                    Text1(Index).Text = ""
'                    MsgBox SQL, vbExclamation
'                    SQL = ""
'                    PonerFoco Text1(Index)
'                End If
'                'Poneos el texto
'                Text2(Index).Text = SQL
'            End If
'        Case 3, 4, 11, 8
'             If Modo = 3 Or Modo = 4 Then
'
'                'Insertando
'                If Text1(Index).Text = "" Then
'                    Text2(Index).Text = "2"
'                    Exit Sub
'                End If
'
'                If Not IsNumeric(Text1(Index).Text) Then
'                    MsgBox "Concepto no es numérico: " & Text1(Index).Text, vbExclamation
'                    Text1(Index).Text = ""
'                    Text2(Index).Text = ""
'                    PonerFoco Text1(Index)
'                    Exit Sub
'                Else
'                    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codConce", Text1(Index).Text, "N")
'                End If
'                If SQL = "" Then
'                    SQL = "Concepto no encontrado: " & Text1(Index).Text
'                    Text1(Index).Text = ""
'                    MsgBox SQL, vbExclamation
'                    PonerFoco Text1(Index)
'                    SQL = ""
'                End If
'                Text2(Index).Text = SQL
'            End If
'        Case 5, 6, 7, 9
'            If Text1(Index).Text = "" Then
'                 Text2(Index).Text = SQL
'                 Exit Sub
'            End If
'            DevfrmCCtas = Text1(Index).Text
'            If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
'                Text1(Index).Text = DevfrmCCtas
'                Text2(Index).Text = SQL
'            Else
'                MsgBox SQL, vbExclamation
'                Text1(Index).Text = ""
'                Text2(Index).Text = ""
'                PonerFoco Text1(Index)
'            End If
'            DevfrmCCtas = ""
'        '....
'    End Select
'    '---
'End Sub
'
'Private Sub HacerBusqueda()
'Dim Cad As String
'Dim CadB As String
'CadB = ObtenerBusqueda(Me)
'
'If chkVistaPrevia = 1 Then
'    MandaBusquedaPrevia CadB
'    Else
'        'Se muestran en el mismo form
'        If CadB <> "" Then
'            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
'            PonerCadenaBusqueda
'        End If
'End If
'End Sub
'
'Private Sub MandaBusquedaPrevia(CadB As String)
'        Dim Cad As String
'        'Llamamos a al form
'        '##A mano
'        Cad = ""
'        Cad = Cad & ParaGrid(Text1(0), 10, "Código")
'        Cad = Cad & ParaGrid(Text1(1), 60, "Denominacion")
'        If Cad <> "" Then
'            Screen.MousePointer = vbHourglass
'            Set frmB = New frmBuscaGrid
'            DevfrmCCtas = ""
'            frmB.vCampos = Cad
'            frmB.vTabla = NombreTabla
'            frmB.vSQL = CadB
'            HaDevueltoDatos = False
'            '###A mano
'            frmB.vDevuelve = "0|1|"
'            frmB.vTitulo = "Formas de pago"
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
'End Sub
'
'
'
'Private Sub PonerCadenaBusqueda()
'Screen.MousePointer = vbHourglass
'On Error GoTo EEPonerBusq
'
'Data1.RecordSource = CadenaConsulta
'Data1.Refresh
'If Data1.Recordset.RecordCount <= 0 Then
'    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'    Else
'        PonerModo 2
'        'Data1.Recordset.MoveLast
'        Data1.Recordset.MoveFirst
'        PonerCampos
'End If
'
'
'Screen.MousePointer = vbDefault
'Exit Sub
'EEPonerBusq:
'    MuestraError Err.Number, "PonerCadenaBusqueda"
'    PonerModo 0
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub PonerCampos()
'    Dim I As Integer
'    Dim mTag As CTag
'    Dim SQL As String
'    If Data1.Recordset.EOF Then Exit Sub
'    PonerCamposForma Me, Data1
'    PonerCtasIVA
'    '-- Esto permanece para saber donde estamos
'    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
'
'End Sub
'
''----------------------------------------------------------------
''----------------------------------------------------------------
''   En PONERMODO se habilitan, o no, los diverso campos del
''   formulario en funcion del modo en k vayamos a trabajar
''
'Private Sub PonerModo(Kmodo As Integer)
'    Dim I As Integer
'    Dim B As Boolean
'    If Modo = 1 Then
'        'Ponemos todos a fondo blanco
'        '### a mano
'        For I = 0 To Text1.Count - 1
'            'Text1(i).BackColor = vbWhite
'            Text1(0).BackColor = &H80000018
'        Next I
'        'chkVistaPrevia.Visible = False
'    End If
'    Modo = Kmodo
'    'chkVistaPrevia.Visible = (Modo = 1)
'
'    'Modo 2. Hay datos y estamos visualizandolos
'    B = (Kmodo = 2)
'    DespalzamientoVisible B
'    'Modificar
'    Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
'    mnModificar.Enabled = B
'    'eliminar
'    Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
'    mnEliminar.Enabled = B
'    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = B
'    Else
'        cmdRegresar.Visible = False
'    End If
'
'    'Modo insertar o modificar
'    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
'    cmdAceptar.Visible = B Or Modo = 1
'    cmdCancelar.Visible = B Or Modo = 1
'    mnOpciones.Enabled = Not B
'    If cmdCancelar.Visible Then
'        cmdCancelar.Cancel = True
'        Else
'        cmdCancelar.Cancel = False
'    End If
'    'Los combo
'    For I = 0 To 3
'        Combo2(I).Enabled = B
'    Next I
'    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 2
'    Toolbar1.Buttons(1).Enabled = Not B
'    Toolbar1.Buttons(2).Enabled = Not B
'
'    If Kmodo = 0 Then lblIndicador.Caption = ""
'
'    '### A mano
'    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
'    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
'    ' Bloqueamos los campos de texto y demas controles en funcion
'    ' del modo en el que estamos.
'    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
'    ' si no  disable. la variable b nos devuelve esas opciones
'    B = (Modo = 2) Or Modo = 0
'    HabilitarText B
'    Combo1.Enabled = Not B
'End Sub
'
'Private Sub HabilitarText(Boleana As Boolean)
'Dim I As Integer
'    On Error Resume Next
'    For I = 0 To Text1.Count - 1
'        Text1(I).Locked = Boleana
'        Text1(I).BackColor = vbWhite
'    Next I
'    For I = 2 To 11
'        imgCuentas(I).Enabled = Not Boleana
'    Next I
'    Err.Clear
'End Sub
'
'
'Private Function DatosOk() As Boolean
'Dim B As Boolean
'    DatosOk = False
'    B = CompForm(Me)
'    If Not B Then Exit Function
'    'Comprobamos  si existe
'    If Modo = 3 Then
'        If DevuelveDesdeBD("codforpa", "sforpa", "codforpa", Text1(0).Text, "N") <> "" Then
'            B = False
'            MsgBox "Ya existe la forma de pago: " & Text1(0).Text, vbExclamation
'        Else
'            B = True
'        End If
'    End If
'    DatosOk = B
'End Function
'
'
''### A mano
''Esto es para que cuando pincha en siguiente le sugerimos
''Se puede comentar todo y asi no hace nada ni da error
''El SQL es propio de cada tabla
'Private Sub SugerirCodigoSiguiente()
'
'    Dim SQL As String
'    Dim RS As ADODB.Recordset
'
'    SQL = "Select Max(codforpa) from " & NombreTabla
'    Text1(0).Text = 1
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, , , adCmdText
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            Text1(0).Text = RS.Fields(0) + 1
'        End If
'    End If
'    RS.Close
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Select Case Button.Index
'Case 1
'    BotonBuscar
'Case 2
'    BotonVerTodos
'Case 6
'    BotonAnyadir
'Case 7
'    If BLOQUEADesdeFormulario(Me) Then BotonModificar
'Case 8
'    BotonEliminar
'Case 12
'    mnSalir_Click
'Case 14 To 17
'    Desplazamiento (Button.Index - 14)
''Case 20
''    'Listado en crystal report
''    Screen.MousePointer = vbHourglass
''    CR1.Connect = Conn
''    CR1.ReportFileName = App.Path & "\Informes\list_Inc.rpt"
''    CR1.WindowTitle = "Listado incidencias."
''    CR1.WindowState = crptMaximized
''    CR1.Action = 1
''    Screen.MousePointer = vbDefault
'
'Case Else
'
'End Select
'End Sub
'
'
'Private Sub DespalzamientoVisible(bol As Boolean)
'    Dim I
'    For I = 14 To 17
'        Toolbar1.Buttons(I).Visible = bol
'    Next I
'End Sub
'
'
'Private Sub PonerCtasIVA()
'Dim SQL As String
'Dim I As Integer
'On Error GoTo EPonerCtasIVA
'
'
''    'Cuentas
''    For I = 5 To 9
''        If I <> 8 Then
''            SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(I).Text, "T")
''            Text2(I).Text = SQL
''        End If
''    Next I
'
'    'Conceptos
'    Text2(3).Text = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(3).Text, "N")
'    Text2(4).Text = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(4).Text, "N")
'    Text2(8).Text = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(8).Text, "N")
'    Text2(11).Text = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(11).Text, "N")
'
'
'    'Diarios
'    Text2(2).Text = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(2).Text, "N")
'    Text2(10).Text = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(10).Text, "N")
'Exit Sub
'EPonerCtasIVA:
'    MuestraError Err.Number, "Poniendo valores ctas. ", Err.Description
'End Sub
'
'
'
'Private Sub PonerFoco(ByRef Text As TextBox)
'    On Error Resume Next
'    Text.SetFocus
'    If Err.Number <> 0 Then Err.Clear
'End Sub
'
'
'
'Private Sub PonerOpcionesMenu()
'    PonerOpcionesMenuGeneral Me
'End Sub
'
'
'
'Private Function SePuedeEliminar() As Boolean
'Dim B As Boolean
'Dim Cad As String
'
'    Screen.MousePointer = vbHourglass
'    SePuedeEliminar = False
'
'    Screen.MousePointer = vbDefault
'End Function
