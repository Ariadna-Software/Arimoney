VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReclama 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de reclamaciones"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   Icon            =   "frmReclama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCta 
      Caption         =   "+"
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   5760
      Width           =   195
   End
   Begin VB.CommandButton cmdObserva 
      Caption         =   "+"
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   6000
      Width           =   195
   End
   Begin VB.TextBox txtaux 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   10
      Tag             =   "observaciones|T|S|||shcocob|observaciones|||"
      Text            =   "frmReclama.frx":000C
      Top             =   6000
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "fecreclama|F|N|||shcocob|fecreclama|dd/mm/yyyy||"
      Text            =   "fecreclama"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   2760
      TabIndex        =   20
      Tag             =   "Cliente|T|N|||shcocob|nommacta|||"
      Text            =   "nommacta"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Tag             =   "codmacta|T|N|||shcocob|codmacta||| "
      Text            =   "codmacta"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   4440
      TabIndex        =   4
      Tag             =   "Importe|N|N|||shcocob|impvenci|###,###,###,##0.00||"
      Text            =   "impvenci"
      Top             =   5775
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   8400
      TabIndex        =   8
      Tag             =   "Vto|N|S|||shcocob|numorden|||"
      Text            =   "numorden"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   7680
      TabIndex        =   7
      Tag             =   "Fec. fact|F|S|||shcocob|fecfaccl|dd/mm/yyyy||"
      Text            =   "fecfaccl"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   6720
      TabIndex        =   6
      Tag             =   "Codigo fra.|N|S|||shcocob|codfaccl|||"
      Text            =   "codfaccl"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "Serie|T|S|||shcocob|numserie|||"
      Text            =   "numserie"
      Top             =   5760
      Width           =   1260
   End
   Begin VB.TextBox txtaux 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   6960
      TabIndex        =   19
      Tag             =   "codigo|N|N|||shcocob|codigo||S|"
      Text            =   "codigo"
      Top             =   5400
      Width           =   1260
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmReclama.frx":001A
      Left            =   2760
      List            =   "frmReclama.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Tipo de Cliente|N|N|||shcocob|carta|||"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10500
      TabIndex        =   11
      Top             =   6180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11700
      TabIndex        =   12
      Top             =   6180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmReclama.frx":001E
      Height          =   5565
      Left            =   60
      TabIndex        =   16
      Top             =   540
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   9816
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
      Left            =   11700
      TabIndex        =   15
      Top             =   6180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Top             =   6015
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   14
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
      TabIndex        =   17
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
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
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   18
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
Attribute VB_Name = "frmReclama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim SQ As String
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
Dim I As Integer

Modo = vModo

B = (Modo = 0)
For I = 0 To txtaux.Count - 2  'El codigo (secuencial)simepre esta oculto
    txtaux(I).Visible = Not B
Next
cmdObserva.Visible = Not B
cmdCta.Visible = Not B
txtaux(I).Visible = False
Combo1.Visible = Not B
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
'If Modo = 2 Then
'   txtaux(0).BackColor = &H80000018
'   Else
'    txtaux(0).BackColor = &H80000005
'End If
'txtaux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = 770
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If
    Limpiar
    txtaux(9).Text = NumF
    
    Combo1.ListIndex = -1
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtaux(0).Text = Format(Now, "dd/mm/yyyy")
    txtaux(0).SetFocus
    
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
    CargaGrid "codigo = -1"
    'Buscar
    Limpiar
    LLamaLineas DataGrid1.Top + 206, 2
    txtaux(0).SetFocus
End Sub


Private Sub Limpiar()
Dim I
    For I = 0 To txtaux.Count - 1
        txtaux(I).Text = ""
    Next
    Me.Combo1.ListIndex = -1
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
    For I = 0 To 9
        If I < 3 Then
            txtaux(I).Text = DataGrid1.Columns(I).Text
        Else
            txtaux(I).Text = DataGrid1.Columns(I + 1).Text
        End If
    Next
    
    NumRegElim = adodc1.Recordset!carta
    Cad = ""
    For I = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(I) = NumRegElim Then
            Combo1.ListIndex = I
            Cad = "OK"
            Exit For
        End If
    Next I
    LLamaLineas anc, 1
   
   'Como es modificar
   txtaux(0).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer
PonerModo xModo + 1
'Fijamos el ancho
For I = 0 To txtaux.Count - 1
    txtaux(I).Top = alto
Next
Combo1.Top = alto - 15
cmdObserva.Top = alto
cmdCta.Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    
    If Not SepuedeBorrar Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar la reclamacion :"
    SQL = SQL & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Cuenta: " & adodc1.Recordset.Fields(1) & " " & DBLet(adodc1.Recordset.Fields(2), "T")
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition - 1
        SQL = "Delete from shcocob where codigo=" & adodc1.Recordset!Codigo
        Conn.Execute SQL
        CargaGrid ""
        
        adodc1.Recordset.Cancel
        If Not adodc1.Recordset.EOF Then
            If adodc1.Recordset.RecordCount >= NumRegElim Then adodc1.Recordset.Move NumRegElim
        End If
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub cmdAceptar_Click()
Dim I As Long
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                espera 0.5
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
                    I = adodc1.Recordset!Codigo
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find ("codigo =" & I)
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
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

Private Sub cmdCta_Click()
         Set frmCCtas = New frmColCtas
         SQ = ""
         frmCCtas.DatosADevolverBusqueda = "0"
         frmCCtas.Show vbModal
         Set frmCCtas = Nothing
         If SQ <> "" Then
            txtaux(1).Text = RecuperaValor(SQ, 1)
            txtaux(2).Text = RecuperaValor(SQ, 2)
            PonerfocoObj Me.Combo1
        End If
End Sub

Private Sub cmdObserva_Click()
    CadenaDesdeOtroForm = txtaux(8).Text
    frmObservaciones.Show vbModal
    txtaux(8).Text = CadenaDesdeOtroForm
    If CadenaDesdeOtroForm <> "" Then PonerfocoObj Me.cmdAceptar
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro a devolver.", vbExclamation
    Exit Sub
End If


Cad = adodc1.Recordset.Fields(0) & "|"
Cad = Cad & adodc1.Recordset.Fields(1) & "|"
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Command1_Click()
    'Ver obseravaciones
    
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then
        cmdRegresar_Click
    Else
        If adodc1.Recordset Is Nothing Then Exit Sub
        If adodc1.Recordset.EOF Then Exit Sub
        
        If Modo = 0 Then
            CadenaDesdeOtroForm = Memo_Leer(adodc1.Recordset!observaciones)
            If CadenaDesdeOtroForm <> "" Then
                frmObservaciones.Show vbModal
                CadenaDesdeOtroForm = ""
            End If
            
        End If
    End If
        
    
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
        '.Buttons(10).Image = 10
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
    CadenaConsulta = "SELECT fecreclama,codmacta,nommacta,if(carta=0,""Carta"",if(carta=1,""Email"",""Teléfono"")) ,impvenci"
    CadenaConsulta = CadenaConsulta & ",numserie,codfaccl,fecfaccl,numorden,observaciones,codigo,carta"
    CadenaConsulta = CadenaConsulta & " FROM shcocob"

    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQ = CadenaSeleccion
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



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    
    SQL = "Select Max(codigo) from shcocob"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            SQL = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = SQL
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
Case 8
        BotonEliminar
Case 11
        frmListado.Opcion = 30
        frmListado.Show vbModal
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
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY fecreclama,fecfaccl,codmacta"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = txtaux(0).Height
    
    
    'Fecha reclama
    I = 0
        DataGrid1.Columns(I).Caption = "F. reclam"
        DataGrid1.Columns(I).Width = 1050
        DataGrid1.Columns(I).NumberFormat = "dd/mm/yyyy"
        
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Cuenta"
        DataGrid1.Columns(I).Width = 1100
        
    
    
    I = 2
        DataGrid1.Columns(I).Caption = "Cliente"
        DataGrid1.Columns(I).Width = 3200
        
    
    
    I = 3
        DataGrid1.Columns(I).Caption = "Envio"
        DataGrid1.Columns(I).Width = 900
            
    'importe
    I = 4
        DataGrid1.Columns(I).Caption = "Importe"
        DataGrid1.Columns(I).Width = 1000
        DataGrid1.Columns(I).NumberFormat = FormatoImporte
        DataGrid1.Columns(I).Alignment = dbgRight
               
    'numserie,codfaccl,numorden,fecfaccl,observaciones,codigo
    I = 5
        DataGrid1.Columns(I).Caption = "Serie"
        DataGrid1.Columns(I).Width = 800
    I = 6
        DataGrid1.Columns(I).Caption = "Nº fra."
        DataGrid1.Columns(I).Width = 1100
        DataGrid1.Columns(I).Alignment = dbgRight
    I = 7
        DataGrid1.Columns(I).Caption = "Fec. fact."
        DataGrid1.Columns(I).Width = 1050
        DataGrid1.Columns(I).NumberFormat = "dd/mm/yyyy"
        DataGrid1.Columns(I).Alignment = dbgRight
    I = 8
        DataGrid1.Columns(I).Caption = "Vto."
        DataGrid1.Columns(I).Width = 500
        DataGrid1.Columns(I).Alignment = dbgCenter
    I = 9
        DataGrid1.Columns(I).Caption = "Obs."
        DataGrid1.Columns(I).Width = 1200
           
      'Codigo NO visible
      DataGrid1.Columns(10).Visible = False
      DataGrid1.Columns(11).Visible = False  'carta
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        J = 60
        For I = 0 To 2
            txtaux(I).Left = DataGrid1.Columns(I).Left + J
            txtaux(I).Width = DataGrid1.Columns(I).Width - 45
        Next
        
        Me.cmdCta.Left = DataGrid1.Columns(2).Left - 45


        
        Combo1.Width = DataGrid1.Columns(3).Width + 120
        Combo1.Left = DataGrid1.Columns(3).Left
        
        For I = 3 To 7
            txtaux(I).Left = DataGrid1.Columns(I + 1).Left + J + J
            txtaux(I).Width = DataGrid1.Columns(I + 1).Width - 60
        Next
        'ajuste manual
        txtaux(3).Width = txtaux(3).Width - 15
        txtaux(4).Left = txtaux(4).Left - 15
        txtaux(5).Width = txtaux(5).Width
        
        
        
        Me.cmdObserva.Left = DataGrid1.Columns(9).Left
        txtaux(8).Left = cmdObserva.Left + 240
        txtaux(8).Width = DataGrid1.Columns(9).Width - 210
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
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

Private Sub txtAux_LostFocus(Index As Integer)
Dim C1 As String
Dim C2 As String
Dim Valor
    txtaux(Index).Text = Trim(txtaux(Index).Text)
    If txtaux(Index).Text = "" Then Exit Sub
    If Modo = 3 Then Exit Sub 'Busquedas

    Select Case Index

    Case 0, 6
        'Fecha
        If Not EsFechaOK(txtaux(Index)) Then
            MsgBox "Fecha incorrecta: " & txtaux(Index).Text, vbExclamation
            txtaux(Index).Text = ""
            Ponerfoco txtaux(Index)
        End If
        
    Case 1
        'Cuenta
        C1 = txtaux(Index).Text
        If CuentaCorrectaUltimoNivel(C1, C2) Then
            txtaux(Index).Text = C1
            If Modo >= 1 Then txtaux(2).Text = C2
            PonerfocoObj Combo1
        Else
            If Modo >= 1 Then
                MsgBox C2, vbExclamation
                txtaux(Index).Text = ""
                Ponerfoco txtaux(Index)
            End If
            
            txtaux(2).Text = ""
            
        End If

    Case 3
        If Not IsNumeric(txtaux(Index).Text) Then
            MsgBox "importe debe ser numérico", vbExclamation
            txtaux(Index).Text = ""
            Ponerfoco txtaux(Index)
        Else
            If InStr(1, txtaux(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtaux(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(txtaux(Index).Text))
            End If
            txtaux(Index).Text = Format(Valor, FormatoImporte)
        End If
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean

B = CompForm(Me)
If Not B Then Exit Function

If Modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", txtaux(0).Text, "N")
     If Datos <> "" Then
        MsgBox "Ya existe la forma de pago : " & txtaux(0).Text & "-" & Datos, vbExclamation
        B = False
    End If
End If
DatosOk = B
End Function

Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open "Select * from stipoformapago order by descformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not miRsAux.EOF
'        Combo1.AddItem miRsAux!descformapago
'        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    Set miRsAux = Nothing
        Combo1.AddItem "Carta"
        Combo1.ItemData(Combo1.NewIndex) = 0
        Combo1.AddItem "Email"
        Combo1.ItemData(Combo1.NewIndex) = 1
        Combo1.AddItem "Teléfono"
        Combo1.ItemData(Combo1.NewIndex) = 2
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False

    
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
