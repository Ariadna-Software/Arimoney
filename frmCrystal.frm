VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCrystal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmCrystal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   4440
      TabIndex        =   10
      Tag             =   "Informe|T|N|||scryst|informe|||"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   2955
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   6000
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Codigo|N|N|0||scryst|codigo||S|"
      Text            =   "Dat"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Tag             =   "Desc.|T|N|||scryst|descripcion|||"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   2955
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   5880
      Width           =   1755
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4200
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCrystal.frx":000C
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   540
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos
Dim Modo As Byte

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

txtAux(0).Visible = Not B
txtAux(1).Visible = Not B
txtAux(2).Visible = Not B
mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(6).Enabled = False 'B And vUsu.Nivel < 2
Me.mnNuevo.Enabled = Toolbar1.Buttons(6).Enabled
Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
Me.mnModificar.Enabled = Toolbar1.Buttons(7).Enabled
Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
Me.mnEliminar.Enabled = Toolbar1.Buttons(8).Enabled



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
txtAux(1).Enabled = (Modo <> 2)
End Sub

'Private Sub BotonAnyadir()
'    Dim NumF As String
'    Dim anc As Single
'
'    'Obtenemos la siguiente numero de factura
'    NumF = SugerirCodigoSiguiente
'    lblIndicador.Caption = "INSERTANDO"
'    'Situamos el grid al final
'    DataGrid1.AllowAddNew = True
'    If adodc1.Recordset.RecordCount > 0 Then
'        DataGrid1.HoldFields
'        adodc1.Recordset.MoveLast
'        DataGrid1.Row = DataGrid1.Row + 1
'    End If
'
'
'
'    If DataGrid1.Row < 0 Then
'        anc = 770
'        Else
'        anc = DataGrid1.RowTop(DataGrid1.Row) + 550
'    End If
'    txtAux(0).Text = NumF
'    txtAux(1).Text = ""
'    LLamaLineas anc, 0
'
'
'    'Ponemos el foco
'    txtAux(0).SetFocus
'
''    If FormularioHijoModificado Then
''        CargaGrid
''        BotonAnyadir
''        Else
''            'cmdCancelar.SetFocus
''            If Not Adodc1.Recordset.EOF Then _
''                Adodc1.Recordset.MoveFirst
''    End If
'End Sub

Private Sub BotonBuscar()
    CargaGrid "codigo= -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    LLamaLineas 770, 2
    txtAux(0).SetFocus
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub



Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 550
    End If
    Cad = ""
    For I = 0 To 1
        Cad = Cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    LLamaLineas anc, 1
   
   'Como es modificar
   txtAux(2).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub


'Private Sub DeseleccionaGrid()
'    On Error GoTo EDeseleccionaGrid
'
'    While Datagrid1.SelBookmarks.Count > 0
'        Datagrid1.SelBookmarks.Remove 0
'    Wend
'    Exit Sub
'EDeseleccionaGrid:
'        Err.Clear
'End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
   ' DeseleccionaGrid DataGrid1
    PonerModo xModo + 1
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
    txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 45 + 15
End Sub

Private Sub BotonEliminar()
    MsgBox "No se pueden eliminar informes", vbCritical
'Dim SQL As String
'    On Error GoTo Error2
'    'Ciertas comprobaciones
'    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
'
'
'    '### a mano
'    SQL = "Seguro que desea eliminar el informe:"
'    SQL = SQL & vbCrLf & "C�digo: " & adodc1.Recordset.Fields(0)
'    SQL = SQL & vbCrLf & "Denominaci�n: " & adodc1.Recordset.Fields(1)
'    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'        'Hay que eliminar
'        SQL = "Delete from scryst where codigo=" & adodc1.Recordset!Codigo
'        Conn.Execute SQL
'        CancelaADODC
'        CargaGrid ""
'        CancelaADODC
'    End If
'Error2:
'        Screen.MousePointer = vbDefault
'        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar informe", Err.Description
End Sub


Private Sub CancelaADODC()
On Error Resume Next
adodc1.Recordset.Cancel
If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Select Case Modo
    Case 1
'    If DatosOk Then
'            '-----------------------------------------
'            'Hacemos insertar
'            If InsertarDesdeForm(Me) Then
'                'MsgBox "Registro insertado.", vbInformation
'                CargaGrid
'                BotonAnyadir
'            End If
'        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CancelaADODC
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
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
    adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ning�n registro devuelto.", vbExclamation
    Exit Sub
End If

    Cad = adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If Modo = 0 Then Unload Me
    End If
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
        .Buttons(10).Image = 16
        .Buttons(11).Image = 15
    End With
    
    
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    CadAncho = False
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    PonerModo 0
    'Cadena consulta
    CadenaConsulta = "Select * from scryst"
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
'BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
'Private Function SugerirCodigoSiguiente() As String
'    Dim SQL As String
'    Dim RS As ADODB.Recordset
'
'    SQL = "Select Max(numdiari) from tiposdiario"
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, , , adCmdText
'    SQL = "1"
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            SQL = CStr(RS.Fields(0) + 1)
'        End If
'    End If
'    RS.Close
'    SugerirCodigoSiguiente = SQL
'End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        'BotonAnyadir
Case 7
        BotonModificar
Case 8
        BotonEliminar
Case 10

'        Screen.MousePointer = vbHourglass
'        If GenerarDiarios Then
'            frmImprimir.NumeroParametros = 0
'            frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
'            frmImprimir.Opcion = 2
'            frmImprimir.Show vbModal
'        End If
Case 11
        Unload Me
Case Else

End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    Dim B As Boolean
    
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codigo"
    lblIndicador.Caption = ""
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "N�"
        DataGrid1.Columns(I).Width = 700
        DataGrid1.Columns(I).NumberFormat = "00"
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Descripci�n"
        DataGrid1.Columns(I).Width = 4000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    I = 2
        DataGrid1.Columns(I).Caption = "Informe"
        DataGrid1.Columns(I).Width = 2200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        txtAux(2).Width = DataGrid1.Columns(2).Width - 60
        CadAncho = True
    End If
   
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
   DataGrid1.Enabled = B
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
'With txtaux(Index)
'    .SelStart = 0
'    .SelLength = Len(.Text)
'End With
 '   PonFoco txtAux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
'    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

txtAux(Index).Text = Trim(txtAux(Index).Text)
If txtAux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas
If Index = 0 Then
    If Not IsNumeric(txtAux(0).Text) Then
        MsgBox "C�digo diario tiene que ser num�rico", vbExclamation
        Exit Sub
    End If
    txtAux(0).Text = Format(txtAux(0).Text, "00")
End If
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

'If Modo = 1 Then
'    'Estamos insertando
'     Datos = DevuelveDesdeBD("numdiari", "tiposdiario", "numdiari", txtAux(0).Text, "T")
'     If Datos <> "" Then
'        MsgBox "Ya existe el diario : " & txtAux(0).Text, vbExclamation
'        B = False
'    End If
'End If
DatosOk = B
End Function


Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False
    SQL = DevuelveDesdeBD("tipoamor", "samort", "numdiari", adodc1.Recordset!numdiari, "N")
    If SQL <> "" Then
        MsgBox "Esta vinculada con parametros de amortizacion", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True
End Function
