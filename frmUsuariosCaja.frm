VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUsuariosCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios cajas"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   Icon            =   "frmUsuariosCaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   290
      Index           =   2
      Left            =   7560
      TabIndex        =   17
      Top             =   5640
      Width           =   135
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   6
      Left            =   7560
      MaxLength       =   30
      TabIndex        =   16
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Diario|N|N|||susucaja|diario|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "saldo|N|N|||susucaja|saldo|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   290
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   5640
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   290
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   5640
      Width           =   135
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   3
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   9
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Cta Caja|T|N|||susucaja|ctacaja|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9540
      TabIndex        =   5
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   7
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Usuario|N|N|0|100|susucaja|codusu|000|S|"
      Text            =   "Dat"
      ToolTipText     =   "1"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmUsuariosCaja.frx":000C
      Height          =   5325
      Left            =   60
      TabIndex        =   13
      Top             =   540
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   9393
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   10
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   11
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
      TabIndex        =   14
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
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
            Object.ToolTipText     =   "PRINCIPAL"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
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
Attribute VB_Name = "frmUsuariosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim CtaDevuelta As String
Dim II As Integer
'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Modo = vModo

B = (Modo = 0)

For II = 0 To 6
    txtAux(II).Visible = Not B
Next II
'txtAux(0).Visible = Not B
'txtAux(1).Visible = Not B
'txtAux(2).Visible = Not B
'txtAux(3).Visible = Not B
Me.Command1(0).Visible = Modo = 1
Me.Command1(1).Visible = Not B
Me.Command1(2).Visible = Not B

mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(8).Enabled = B
Toolbar1.Buttons(7).Enabled = B
Toolbar1.Buttons(6).Enabled = B

'Prueba
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtAux(0).BackColor = &H80000018
   Else
    txtAux(0).BackColor = &H80000005
End If
txtAux(0).Enabled = (Modo <> 2)
End Sub


Private Function TienePermiso() As Boolean
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene permisos", vbExclamation
        TienePermiso = False
    Else
        TienePermiso = True
    End If
End Function

Private Sub BotonAnyadir()

    Dim anc As Single
    
    If Not TienePermiso Then Exit Sub
    
    
    'Obtenemos la siguiente numero de factura
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = 750
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If
    For II = 0 To 6
        txtAux(II).Text = ""
    Next II
    'txtaux(1).Text = ""
    'txtaux(2).Text = ""
    'txtaux(3).Text = ""
    'Combo1.ListIndex = -1
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "ctacaja = -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    LLamaLineas DataGrid1.Top + 206, 2
    txtAux(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If Not TienePermiso Then Exit Sub

    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    For I = 0 To 6
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next
    
    LLamaLineas anc, 1
   
   'Como es modificar
   txtAux(2).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)

PonerModo xModo + 1
'Fijamos el ancho
For II = 0 To 6
    If II < 3 Then Me.Command1(II).Top = alto
    txtAux(II).Top = alto
Next II
'txtaux(0).Top = alto
'txtaux(1).Top = alto
'txtaux(2).Top = alto
'txtaux(3).Top = alto
'Me.Command1(0).Top = alto
'Me.Command1(1).Top = alto
'Me.Command1(2).Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub

    If Not TienePermiso Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar el usuario/caja:"
    SQL = SQL & vbCrLf & "Nombre: " & Adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "caja: " & Adodc1.Recordset.Fields(2) & " - " & Adodc1.Recordset.Fields(3)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        
        'Hay que eliminar
        SQL = "Delete from susucaja where codusu=" & Adodc1.Recordset.Fields(0)
        Conn.Execute SQL
        CargaGrid ""
        Adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    I = Adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If Adodc1.Recordset.EOF Then
    MsgBox "Ning�n registro a devolver.", vbExclamation
    Exit Sub
End If


If Adodc1.Recordset.Fields(0) >= 900 Then
    If vUsu.Nivel > 1 Then
        MsgBox "Los conceptos superiores a 900 se los reserva la aplicaci�n.", vbExclamation
        Exit Sub
    Else
        Cad = "Los conceptos superiores a 900 son de la aplicaci�n y no deberia utilizarlis. �Desea continuar de igual modo?"
        If MsgBox(Cad, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If
End If
Cad = Adodc1.Recordset.Fields(0) & "|"
Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
Cad = Cad & Adodc1.Recordset.Fields(2) & "|"
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Command1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass

    If Index = 1 Then
        'CUENTA CAJA
        CtaDevuelta = ""
        Set frmCta = New frmColCtas
        frmCta.DatosADevolverBusqueda = "0|1"
        frmCta.ConfigurarBalances = 3
        frmCta.Show vbModal
        Set frmCta = Nothing
        If CtaDevuelta <> "" Then
            Me.Refresh
            txtAux(2).Text = CtaDevuelta
            txtAux_LostFocus 2
        End If
    Else
        Set frmB = New frmBuscaGrid
        If Index = 0 Then

            CtaDevuelta = "C�digo|codusu|N|30�"
            CtaDevuelta = CtaDevuelta & "Nombre|nomusu|T|60�"
            
            frmB.vTitulo = "Usuarios sistema contabilidad"
            frmB.vTabla = "Usuarios.Usuarios"
            frmB.vSQL = "  nivelusu>0"
            II = 0  'el txtaux(0)
        Else
            '2: tipos diario
            CtaDevuelta = "C�digo|numdiari|N|30�"
            CtaDevuelta = CtaDevuelta & "Descripcion|desdiari|T|60�"
            
            frmB.vTitulo = "DIARIOS"
            frmB.vTabla = "tiposdiario"
            frmB.vSQL = ""
            II = 5
        End If
        '###A mano
        frmB.vCampos = CtaDevuelta
        frmB.vDevuelve = "0|1|"
        
        frmB.vSelElem = 0
        '#
        CtaDevuelta = ""
        frmB.Show vbModal
        Set frmB = Nothing
        If CtaDevuelta <> "" Then
            
            txtAux(II).Text = CtaDevuelta
            txtAux_LostFocus II
                
        End If
    End If
End Sub

Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 18
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    Me.Icon = frmPpal.Icon
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    CargaCombo
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    'Consulta compleja, compleja
    
    
'      ANTES 2.0.9
'    CadenaConsulta = "Select usucaja.codusu,u1.nomusu, ctacaja, nommacta , if(predeterminado=0,"""","" * "") as Ppal"
'    CadenaConsulta = CadenaConsulta & " from usucaja,usuarios.usuarios as u1,cuentas "
'    CadenaConsulta = CadenaConsulta & " WHERE usucaja.codusu = u1.codusu and ctacaja = cuentas.codmacta"
'
    
     

  
    CadenaConsulta = "Select susucaja.codusu,u1.nomusu, ctacaja, nommacta , saldo,diario,desdiari,if(predeterminado=0,"""","" * "") as Ppal"
    CadenaConsulta = CadenaConsulta & " from susucaja,usuarios.usuarios as u1,cuentas,tiposdiario"
    CadenaConsulta = CadenaConsulta & " WHERE susucaja.codusu = u1.codusu and ctacaja = cuentas.codmacta and tiposdiario.numdiari=susucaja.diario"
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CtaDevuelta = RecuperaValor(CadenaDevuelta, 1)
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    CtaDevuelta = RecuperaValor(CadenaSeleccion, 1)
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
Case 8
        
        BotonEliminar
Case 10
        If vUsu.Nivel > 1 Then
            MsgBox "No esta autorizado", vbExclamation
            Exit Sub
        End If
        
        NumRegElim = 1
        If Not Adodc1.Recordset Is Nothing Then
            If Not Adodc1.Recordset.EOF Then NumRegElim = 0
        End If
        
        If NumRegElim = 1 Then
            MsgBox "No existen , o no se han cargado los datos", vbExclamation
            Exit Sub
        End If
        CtaDevuelta = "Va a cambiar a caja principal predterminada a : " & vbCrLf & _
            Adodc1.Recordset!nomusu & " - " & Adodc1.Recordset!Nommacta
        CtaDevuelta = CtaDevuelta & vbCrLf & vbCrLf & "�Desea continuar?"
        If MsgBox(CtaDevuelta, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        
        'Predeterminamos
        CtaDevuelta = "UPDATE susucaja set predeterminado=0"
        Conn.Execute CtaDevuelta
        
        CtaDevuelta = "UPDATE susucaja SET predeterminado=1 WHERE codusu =" & Adodc1.Recordset!codusu
        CtaDevuelta = CtaDevuelta & " AND ctacaja = '" & Adodc1.Recordset!CtaCaja & "'"
        Conn.Execute CtaDevuelta
        CargaGrid
Case 11

Case 12
        Unload Me
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub

Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    

    
    Adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codusu"
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    'Set DataGrid1.DataSource = adodc1
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Cod."
        DataGrid1.Columns(I).Width = 500
        DataGrid1.Columns(I).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Nombre"
        DataGrid1.Columns(I).Width = 2000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'El importe es campo calculado
    I = 2
        DataGrid1.Columns(I).Caption = "Cta caja"
        DataGrid1.Columns(I).Width = 1100
    I = 3
        DataGrid1.Columns(I).Caption = "Denominaci�n"
        DataGrid1.Columns(I).Width = 2250
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    
    I = 4
        DataGrid1.Columns(I).Width = 1000
        DataGrid1.Columns(I).NumberFormat = FormatoImporte
        DataGrid1.Columns(I).Alignment = dbgRight
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    I = 5
        DataGrid1.Columns(I).Caption = "Diario"
        DataGrid1.Columns(I).Width = 600
        DataGrid1.Columns(I).Alignment = dbgRight
    I = 6
        DataGrid1.Columns(I).Caption = "Desc. diario"
        DataGrid1.Columns(I).Width = 1800
    
    
    
        'PPAL
    I = 7
        DataGrid1.Columns(I).Width = 600
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    
    
    
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        'txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        'txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        'txtAux(2).Width = DataGrid1.Columns(2).Width - 60
        'txtAux(3).Width = DataGrid1.Columns(3).Width - 60
        'txtAux(3).Width = DataGrid1.Columns(3).Width - 60
        
        For I = 0 To 6
            txtAux(I).Width = DataGrid1.Columns(I).Width - 60
        Next I
        
        
        'Combo1.Width = DataGrid1.Columns(3).Width
        txtAux(0).Left = DataGrid1.Left + 340
        Me.Command1(0).Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 45
        Me.Command1(1).Left = txtAux(2).Left + txtAux(2).Width
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 90
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 75
        txtAux(5).Left = txtAux(4).Left + txtAux(4).Width + 45
        Me.Command1(2).Left = txtAux(5).Left + txtAux(5).Width
        txtAux(6).Left = txtAux(5).Left + txtAux(5).Width + 60
        
        
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String
Dim C2 As String

    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then
        Select Case Index
        Case 0
            II = 1
        Case 2
            II = 2
        Case 5
            II = 6
        Case Else
            II = -1
        End Select
        If II > 0 Then txtAux(II).Text = ""
        Exit Sub
    End If
    If Modo = 3 Then Exit Sub 'Busquedas
    Select Case Index
    Case 0
        If Not IsNumeric(txtAux(0).Text) Then
            MsgBox "C�digo usuario tiene que ser num�rico", vbExclamation
            Cad = "MAL"
        Else
            Cad = ""
        End If
        
        If Cad = "" Then
            Cad = "nivelusu"
            C2 = DevuelveDesdeBD("nomusu", "usuarios.usuarios", "codusu", txtAux(0).Text, "N", Cad)
            If C2 <> "" Then
                If Val(Cad) < 0 Then C2 = ""
            End If
        
        End If
        
        
        txtAux(1).Text = C2
        
        txtAux(0).Text = Format(txtAux(0).Text, "000")
    Case 2
              C2 = txtAux(2).Text
              If CuentaCorrectaUltimoNivel(C2, Cad) Then
                
              Else
                C2 = ""
                Cad = ""
              End If
              txtAux(3).Text = Cad
              txtAux(2).Text = C2
    Case 4
        If Not IsNumeric(txtAux(4).Text) Then
            MsgBox "Campo num�rico", vbExclamation
        Else
            If InStr(1, txtAux(4).Text, ",") Then
                txtAux(4).Text = Format(ImporteFormateado(txtAux(4).Text), FormatoImporte)
            Else
                txtAux(4).Text = Format(Format(TransformaPuntosComas(txtAux(4).Text), FormatoImporte), FormatoImporte)
            End If
        End If
    Case 5
        C2 = ""
        If Not IsNumeric(txtAux(5).Text) Then
            Cad = "Diario tiene que ser num�rico"
        Else
            Cad = ""
        End If
        
        If Cad = "" Then
            C2 = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtAux(5).Text, "N")
            If C2 = "" Then Cad = "Diario tiene que ser num�rico"
        End If
        If Cad <> "" Then MsgBox Cad, vbInformation
        txtAux(6).Text = C2
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean

B = CompForm(Me)
If Not B Then Exit Function

If CuentaBloqeada(Me.txtAux(1).Text, Now, True) Then Exit Function


DatosOk = B
End Function

Private Sub CargaCombo()
'    Combo1.Clear
'    'Conceptos
'    Combo1.AddItem "Debe"
'    Combo1.ItemData(Combo1.NewIndex) = 1
'
'    Combo1.AddItem "Haber"
'    Combo1.ItemData(Combo1.NewIndex) = 2
'
'    Combo1.AddItem "Decide asiento"
'    Combo1.ItemData(Combo1.NewIndex) = 3
End Sub


Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
'Dim SQL As String
    SepuedeBorrar = False
    'Comprobamos k no hay en caja nada por traspasar
    CtaDevuelta = DevuelveDesdeBD("feccaja", "scacaja", "codusu", Adodc1.Recordset!codusu, "T")
    If CtaDevuelta <> "" Then
        MsgBox "Tiene datos en caja pendientes de llevar a contabilidad", vbExclamation
        Exit Function
    End If
'    CtaDevuelta = DevuelveDesdeBD("ctacaja", "shcaja", "ctacaja", Adodc1.Recordset!CtaCaja, "T")
'    If CtaDevuelta <> "" Then
'        MsgBox "Hay datos en el historico de caja pendiente de elminar.", vbExclamation
'        Exit Function
'    End If
    CtaDevuelta = ""
    SepuedeBorrar = True
End Function


Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub



Private Sub KEYpress(KeyAscii As Integer)
    'Caption = KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            If Modo = 0 Then Unload Me
        End If
    End If
End Sub

