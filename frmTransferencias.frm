VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransferencias2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "frmTransferencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboConcepto 
      Height          =   315
      ItemData        =   "frmTransferencias.frx":000C
      Left            =   5040
      List            =   "frmTransferencias.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Usuario|N|N|||stransfer|conceptoTrans|||"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   3
      Left            =   4680
      TabIndex        =   14
      Tag             =   "Usuario|T|N|||@@@|codmacta|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10080
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   5
      Left            =   6840
      TabIndex        =   13
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   4
      Left            =   5400
      TabIndex        =   4
      Tag             =   "Usuario|T|N|||stransfer|codmacta|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||@@@|Fecha|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10620
      TabIndex        =   6
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9420
      TabIndex        =   5
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
      Tag             =   "Codigo|N|N|0||@@@|codigo|000|S|"
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
      Tag             =   "Descripción|T|N|||@@@|Descripcion|||"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   90
      TabIndex        =   8
      Top             =   6000
      Width           =   1755
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Object.ToolTipText     =   "Modificar vtos"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar diskette"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Contabilizar"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   10
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
      Bindings        =   "frmTransferencias.frx":0039
      Height          =   5295
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   11535
      _ExtentX        =   20346
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   6120
      Width           =   7095
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
Attribute VB_Name = "frmTransferencias2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'0. Transferencias ABONOS(cobros)    1.- Transferencias PAGOS   2.- Pagos domiciliados
Public TipoDeFrm As Byte
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Tabla As String
Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos
Dim Modo As Byte

Private NIF As String

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
    txtAux(3).Visible = Not B
    txtAux(4).Visible = Not B
    txtAux(5).Visible = Not B
    Me.cboConcepto.Visible = Not B 'And TipoDeFrm < 2
    mnOpciones.Enabled = B
    Toolbar1.Buttons(1).Enabled = B
    Toolbar1.Buttons(2).Enabled = B
    Toolbar1.Buttons(6).Enabled = B
    Toolbar1.Buttons(7).Enabled = B
    Toolbar1.Buttons(8).Enabled = B
    Toolbar1.Buttons(10).Enabled = B
    Toolbar1.Buttons(11).Enabled = B
    Toolbar1.Buttons(12).Enabled = B
    
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

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    NumF = DevuelveDesdeBD("nifempre", "empresa2", "1", "1", "N")
    If NumF = "" Then
        MsgBox "La empresa no tiene NIF. No puede generar fichero norma bancaria", vbExclamation
        Exit Sub
    Else
        NIF = NumF
    End If
    
    
    
    MOntaSQL2 NumF, 0
    If Not VerHayEfectos(NumF) Then
        NIF = "una nueva transferencia"
        If Me.TipoDeFrm = 2 Then
            If vParam.PagosConfirmingCaixa Then
                'Pica
                NIF = "una nueva remesa CaixaConfirming"
            Else
                NIF = "un nuevo pago domiciliado"
            End If
        End If
        
        MsgBox "No hay efectos para realizar " & NIF, vbExclamation
        NIF = ""
        Exit Sub
        'abono
        If TipoDeFrm = 0 Then
            'Puede ser que existan cobros con importe negativo, pero con forma de pago
            'distinta de transferencia
            'Con esta pregunta podriamos pasarlos a forma de pago=transferencia were impvenci<0
        
        End If
    End If
    
    DatosADevolverBusqueda = ""
    frmVarios.Opcion = 15
    frmVarios.SubTipo = TipoDeFrm
    frmVarios.Show vbModal
    
    
    If DatosADevolverBusqueda <> "" Then
        'Recargamos el datagrid
        'Situamos el ado donde toca
        'Lanzamos la generacion del diskette
        CargaGrid ""
        If SituarData(Me.adodc1, " codigo =" & DatosADevolverBusqueda, "") Then
              'LANZAMOS LA GENERACION DEL DISKETTE
                If GeneraNormaBancaria() Then
                    'UPDATEAMOS LA TABLA DE transferencias poniendo
                    'la marca de llevado a banco
                    
                End If
        Else
            MsgBox "Error situando el Recordset: " & DatosADevolverBusqueda, vbExclamation
        End If
    End If


    
End Sub

Private Sub BotonBuscar()
    CargaGrid " codigo = -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    Me.txtAux(5).Text = ""
    Me.cboConcepto.ListIndex = -1
    LLamaLineas 820, 2
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
    DeseleccionaGrid DataGrid1
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 615
    End If
    Cad = ""
    For I = 0 To 1
        Cad = Cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    For I = 0 To 3
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    For I = 4 To 5
        txtAux(I).Text = DataGrid1.Columns(I + 2).Text
    Next I
    I = DBLet(Me.adodc1.Recordset!conceptoTrans, "N")
    
    If TipoDeFrm = 2 Then
        Me.cboConcepto.ListIndex = I
    Else
        'nomina-pension-ordinaria  (1-9-0)
        If I = 9 Then
            Me.cboConcepto.ListIndex = 1
        Else
            If I = 1 Then
                Me.cboConcepto.ListIndex = 0
            Else
                Me.cboConcepto.ListIndex = 2
            End If
        End If
    End If
    LLamaLineas anc, 1
   
   'Como es modificar
   txtAux(1).SetFocus
   
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
Dim I As Integer
   ' DeseleccionaGrid DataGrid1
    PonerModo xModo + 1
    'Fijamos el ancho
    For I = 0 To 5
        txtAux(I).Top = alto
    Next I
    Me.cboConcepto.Top = alto
    txtAux(0).Left = DataGrid1.Left + 340
    For I = 1 To 2
        txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 45
    Next I
    txtAux(3).Left = DataGrid1.Columns(4).Left '+ 60
    txtAux(3).Width = DataGrid1.Columns(4).Left '+ 60
    cboConcepto.Left = DataGrid1.Columns(5).Left + 145
    txtAux(4).Left = cboConcepto.Left + cboConcepto.Width + 60
    txtAux(5).Left = txtAux(4).Left + txtAux(4).Width + 45
    
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    If Not BloqueoManual(True, "Transferencias", CStr(vEmpresa.codempre)) Then
        MsgBox "El proceso esta bloqueado por otro usuario", vbExclamation
        Exit Sub
     End If
    '### a mano
    SQL = "Seguro que desea eliminar :"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descripcion: " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(2)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        If TipoDeFrm <> 0 Then
            SQL = "UPDATE spagop set transfer=NULL "
            'Abril 2015
            SQL = SQL & ", emitdocum =0"
            SQL = SQL & " where transfer =" & adodc1.Recordset!Codigo
        Else
            SQL = "UPDATE scobro set transfer=NULL where transfer =" & adodc1.Recordset!Codigo
        End If
        
        Conn.Execute SQL
        If TipoDeFrm <> 0 Then
            SQL = "Delete from stransfer where codigo=" & adodc1.Recordset!Codigo
        Else
            SQL = "Delete from stransfercob where codigo=" & adodc1.Recordset!Codigo
        End If
        
        Conn.Execute SQL
        CancelaADODC
        CargaGrid ""
        CancelaADODC
    End If
    BloqueoManual False, "Transferencias", ""
        
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar en " & Me.Caption, Err.Description
End Sub


Private Sub CancelaADODC()
On Error Resume Next
adodc1.Recordset.Cancel
If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cboConcepto_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
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
                
                                
                ' BLOQUEAR
               
                PonerModo 0
                DataGrid1.AllowAddNew = False
                CargaGrid
                
                If SituarData(Me.adodc1, "codigo = " & txtAux(0).Text, txtAux(0).Text) Then
                    'Lanzamos
                     If BloqueoManual(True, "transfer", adodc1.Recordset!Codigo) Then
                     
                        'Lanzamos la pantalla para cargar datos
                        NumRegElim = adodc1.Recordset!Codigo
                     
                        MostrarPantallaConVencimientos
                        

                        
                     
                     
                     
                     
                     
                     
                        'Si al volver NUMREGELIM vale <=0
                        'Entonces no se ha añaido ningun registro
                        If NumRegElim = 0 Then
                            MsgBox "Deberias borrar esta transferencia. No tienen efectos.", vbExclamation
                        Else
                            'LANZAMOS LA GENERACION DEL DISKETTE
                            If GeneraNormaBancaria() Then
                                'UPDATEAMOS LA TABLA DE transferencias poniendo
                                'la marca de llevado a banco
                                
                            End If
                        End If
                     End If
                End If
                BloqueoManual False, "transfer", ""
                
            End If
        End If
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

    lblIndicador.Caption = ""
End Sub

Public Function VerHayEfectos(ByVal miSql As String) As Boolean
    Set miRsAux = New ADODB.Recordset
    VerHayEfectos = False
    If TipoDeFrm <> 0 Then
        miSql = " FROM spagop,sforpa WHERE spagop.codforpa = sforpa.codforpa AND " & miSql
    Else
       miSql = " FROM scobro,sforpa WHERE scobro.codforpa = sforpa.codforpa AND " & miSql
       miSql = miSql & " and impvenci<0 "
    End If
    miRsAux.Open "SELECT Count(*) " & miSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then VerHayEfectos = True
    End If
    miRsAux.Close
    Set miRsAux = Nothing

End Function
'vnuevo: 0.- Nuevo, 1-modificar, 2 contabilizar
Private Sub MOntaSQL2(ByRef vSQL As String, vnuevo As Byte)
    
    If TipoDeFrm <> 0 Then
        
        If TipoDeFrm = 1 Then
            vSQL = "1" '1: TRANSFERENCIAS
        Else
            vSQL = "5" '5: CONFIRMING en pagos=PAGO domiciliado  o CaixaConfirming
        End If
        vSQL = " sforpa.tipforpa = " & vSQL
        If vnuevo = 0 Then
            vSQL = vSQL & " AND spagop.transfer is null and impefect - coalesce(imppagad,0)>0"
        Else
            If vnuevo = 1 Then
                vSQL = vSQL & " AND ((spagop.transfer is null) or spagop.transfer = " & adodc1.Recordset!Codigo & ") and impefect > 0"
            Else
                vSQL = vSQL & " AND spagop.transfer = " & adodc1.Recordset!Codigo
            End If
        End If
        
        
    Else
        'ABONOS en copbros
        vSQL = " impvenci <0 "
        If vnuevo = 0 Then
            vSQL = vSQL & " AND scobro.transfer is null"
        Else
            If vnuevo = 1 Then
                vSQL = vSQL & " AND ((scobro.transfer is null) or scobro.transfer = " & adodc1.Recordset!Codigo & ")"
            Else
                vSQL = vSQL & " AND scobro.transfer = " & adodc1.Recordset!Codigo
            End If
        End If
        vSQL = "(" & vSQL & ")"
    End If
End Sub


Private Sub MostrarPantallaConVencimientos()
Dim SQL As String
Dim Cad As String


    Screen.MousePointer = vbHourglass
    
    'Hacemos un conteo
    MOntaSQL2 SQL, 0
    
    If Not VerHayEfectos(SQL) Then
        MsgBox "Ningún dato con esos valores.", vbExclamation
    Else
        'Hay datos, abriremos el forumalrio para k seleccione
        'los pagos que queremos hacer
        With frmVerCobrosPagos
            .vSQL = SQL
            .OrdenarEfecto = True
            .Regresar = False
            .Cobros = False
            'Los texots
            .Tipo = 1
            '.vTextos = Text1(5).Text & "|" & Me.txtCta(1).Text & " - " & Me.txtDescCta(1).Text & "|" & SubTipo & "|"
            .vTextos = adodc1.Recordset!Fecha & "|" & adodc1.Recordset!codmacta & " - " & adodc1.Recordset!Nommacta & "|1|  '1: transferencia"
            .SegundoParametro = NumRegElim
            
            NumRegElim = 0
            
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
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

Private Sub cmdRegresar_Click()
Dim Cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
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
        
        
        .Buttons(10).Image = 10
        .Buttons(11).Image = 28
        .Buttons(12).Image = 25
        
        .Buttons(14).Image = 16
        .Buttons(15).Image = 15
    End With
    
    Tabla = "stransfer"
    Caption = "Transferencias"
'    If TipoDeFrm = 0 Then
'        Tabla = Tabla & "cob"
'        Caption = Caption & " COBROS (Abonos)"
'    Else
'        If TipoDeFrm = 1 Then
'            Caption = Caption & " PAGOS "
'        Else
'            Caption = "Pagos domiciliados"
'        End If
'    End If
    Label1.Caption = ""
    Select Case TipoDeFrm
    Case 0
        Tabla = Tabla & "cob"
        Caption = Caption & " COBROS (Abonos)"
    Case 1
        Caption = Caption & " PAGOS "
    Case 2
        If vParam.PagosConfirmingCaixa Then
            Caption = "Caixa confirming"
            Label1.Caption = "Caixa confirming"
        Else
            Caption = "Pagos domiciliados"
            Label1.Caption = " P A G O S     D O M I C I L I A D O S"
        End If
    End Select
    
    'Segun sea una tabla u otra, los tagas apuntaran a una u otra ;)
    For Modo = 0 To 3
        Me.txtAux(Modo).Tag = Replace(Me.txtAux(Modo).Tag, "@@@", Tabla)
    Next
    
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    CadAncho = False
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    PonerModo 0
    Label1.Visible = Me.TipoDeFrm = 2
    
    'Cadena consulta
    CadenaConsulta = "Select codigo, Descripcion, fecha,importe,conceptoTrans,"
    
    
    'Nominas
        ' 0.- pension     1.- Abono nomina     '9.- Transferencia ordinaria
    'Pagos domiciliados
        ' 0.- Al vencimiento    1.- A una fecha
    If TipoDeFrm < 2 Then
        'transfer
        CadenaConsulta = CadenaConsulta & "if(conceptoTrans=9,'Ordinaria',if(conceptoTrans=1,'Nómina','Pensión')),"
    Else
        'pagos dom
        CadenaConsulta = CadenaConsulta & "if(conceptoTrans=0,'Vencimiento','Fecha intro.'),"
    End If
    CadenaConsulta = CadenaConsulta & Tabla & ".codmacta,nommacta from " & Tabla & ",cuentas"
    CadenaConsulta = CadenaConsulta & " where cuentas.codmacta = " & Tabla & ".codmacta"
    'Pagos
    If Me.TipoDeFrm > 0 Then
        ' '0-transfere 1 pago DOMiciliado       2.- Confirming
        CadenaConsulta = CadenaConsulta & " AND subtipo = " & Me.TipoDeFrm - 1
    End If
    CargaGrid
    
    'Combo.
    Me.cboConcepto.Clear
    If TipoDeFrm = 2 Then
        cboConcepto.AddItem "Vencimiento"
        cboConcepto.ItemData(cboConcepto.NewIndex) = 0
        cboConcepto.AddItem "Fecha intro."
        cboConcepto.ItemData(cboConcepto.NewIndex) = 1
    Else
        cboConcepto.AddItem "Nómina"
        cboConcepto.ItemData(cboConcepto.NewIndex) = 1
        cboConcepto.AddItem "Ordinaria"
        cboConcepto.ItemData(cboConcepto.NewIndex) = 9
        cboConcepto.AddItem "Pension"
        cboConcepto.ItemData(cboConcepto.NewIndex) = 0
    End If
    
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
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As Long
Dim N As Long
Dim AUX As String
Dim Cad As String


    
    
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
        
        
Case 10, 12
     If adodc1.Recordset.EOF Then Exit Sub
     
     
     NIF = ""
     If adodc1.Recordset!Fecha < vParam.fechaAmbito Then
        NIF = "Fecha fuera de ambito"
    Else
        If adodc1.Recordset!Fecha > DateAdd("yyyy", 1, vParam.fechafin) Then NIF = "Fecha mayor fin ejercicio siguiente"
    End If
    
    If adodc1.Recordset!Importe = 0 Then NIF = NIF & vbCrLf & " Transferencia ya contabilizada"

    If NIF <> "" Then
        MsgBox NIF, vbExclamation
        Exit Sub
    End If
     'Hacemos un conteo
     NIF = "TRANSFERENCIAS" & TipoDeFrm
    
        
     If Not BloqueoManual(True, NIF, CStr(vEmpresa.codempre)) Then
        MsgBox "El proceso esta bloqueado por otro usuario", vbExclamation
        Exit Sub
     End If
     If Button.Index = 10 Then
        MOntaSQL2 NIF, 1
     Else
         MOntaSQL2 NIF, 2
     End If
     With frmVerCobrosPagos
            .vSQL = NIF
            .OrdenarEfecto = True
            .Regresar = False
            .Cobros = TipoDeFrm = 0
            .ContabTransfer = (Button.Index = 12)
            'Los texots
            .Tipo = 1
            '.vTextos = Text1(5).Text & "|" & Me.txtCta(1).Text & " - " & Me.txtDescCta(1).Text & "|" & SubTipo & "|"
            .vTextos = adodc1.Recordset!Fecha & "|" & adodc1.Recordset!codmacta & " - " & adodc1.Recordset!Nommacta & "|"  'antes ponia aqui: |1|
            If Me.TipoDeFrm = 2 Then
                'Es un pago domiciliado
                .vTextos = .vTextos & "5||PAGO DOMICILIADO|"
            Else
                .vTextos = .vTextos & "1||"
            End If
            .SegundoParametro = adodc1.Recordset!Codigo
            NumRegElim = 0
            .Show vbModal
            
    End With
    Screen.MousePointer = vbHourglass
    'UPDATEAMOS EL IMPORTE
    If TipoDeFrm = 0 Then
        NIF = "Select sum(impvenci) from scobro"
    Else
        NIF = "Select sum(impefect) from spagop "
    End If
    NIF = NIF & " WHERE transfer = " & adodc1.Recordset!Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open NIF, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NIF = "UPDATE stransfer"
    If TipoDeFrm = 0 Then NIF = NIF & "cob"
    
    NIF = NIF & " SET importe =" & TransformaComasPuntos(CStr(Abs(DBLet(miRsAux.Fields(0), "N"))))
    miRsAux.Close
    Set miRsAux = Nothing
    NIF = NIF & " WHERE codigo =" & adodc1.Recordset!Codigo
    Ejecuta NIF
    espera 0.3
    NIF = "TRANSFERENCIAS" & TipoDeFrm
    BloqueoManual False, NIF, ""
    CargaGrid ""
    Screen.MousePointer = vbDefault
Case 11
    If Not adodc1.Recordset.EOF Then
            NIF = DevuelveDesdeBD("nifempre", "empresa2", "codigo", "1", "N")
            If NIF = "" Then
                MsgBox "La empresa no tiene NIF. No puede generar transferencias o pagos domiciliados", vbExclamation
                Exit Sub
            End If
            
            GeneraNormaBancaria
    End If

Case 14
    '0. Transferencias ABONOS(cobros)    1.- Transferencias PAGOS
    '  2.- Pagos domiciliados (pueden ser caixa confirming
    Select Case TipoDeFrm
    Case 0
        frmListado.Opcion = 13
    Case 1
        frmListado.Opcion = 11
    Case 2
        If vParam.PagosConfirmingCaixa Then
            frmListado.Opcion = 44
        Else
            frmListado.Opcion = 43
        End If
    End Select
    frmListado.Show vbModal
Case 15
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
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY fecha desc,codigo desc"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
      adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Nº"
        DataGrid1.Columns(I).Width = 650
        DataGrid1.Columns(I).NumberFormat = "00"
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Descripcion"
        DataGrid1.Columns(I).Width = 3100
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    I = 2
        DataGrid1.Columns(I).Caption = "Fecha"
        DataGrid1.Columns(I).Width = 1000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    I = 3
        DataGrid1.Columns(I).Caption = "Importe"
        DataGrid1.Columns(I).Width = 1100
        DataGrid1.Columns(I).Alignment = dbgRight
        DataGrid1.Columns(I).NumberFormat = FormatoImporte
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    I = 4
        DataGrid1.Columns(I).Caption = ""
        DataGrid1.Columns(I).Width = 0
            
            
    I = 5
        If TipoDeFrm < 2 Then
            DataGrid1.Columns(I).Caption = "Concepto"
        Else
            DataGrid1.Columns(I).Caption = "Fec. pago"
        End If
        DataGrid1.Columns(I).Width = 1050
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
            
    I = 6
        DataGrid1.Columns(I).Caption = "Cuenta"
        DataGrid1.Columns(I).Width = 1050
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
                    
    I = 7
        DataGrid1.Columns(I).Caption = "Descripcion"
        DataGrid1.Columns(I).Width = 2800
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        For I = 0 To 3
            txtAux(I).Width = DataGrid1.Columns(I).Width - 45
        Next I
        Me.cboConcepto.Width = DataGrid1.Columns(5).Width - 45
        For I = 4 To 5
            txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 45
        Next I
        

        CadAncho = True
    End If
   
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(10).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(11).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(12).Enabled = Not adodc1.Recordset.EOF
   DataGrid1.Enabled = B
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
'With txtaux(Index)
'    .SelStart = 0
'    .SelLength = Len(.Text)
'End With
    'Ponerfoco txtAux(Index)
    ObtenerFoco txtAux(Index)
End Sub

Private Sub ObtenerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
    KeyPressGral KeyAscii
    
    
End Sub




Private Sub txtAux_LostFocus(Index As Integer)
Dim C2 As String
Dim Cad As String
txtAux(Index).Text = Trim(txtAux(Index).Text)
If txtAux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas
Select Case Index
Case 0
    If Not IsNumeric(txtAux(0).Text) Then
        MsgBox "Código diario tiene que ser numérico", vbExclamation
        Exit Sub
    End If
Case 2
    
        If Not EsFechaOK(txtAux(2)) Then Ponerfoco txtAux(2)
Case 3
        
        C2 = txtAux(3).Text
        If CuentaCorrectaUltimoNivel(C2, Cad) Then
            C2 = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", C2, "T")
            If C2 = "" Then
                MsgBox "la cuenta no esta asociada a ningun banco", vbExclamation
                Cad = ""
            End If
        Else
            C2 = ""
            Cad = ""
        End If
        txtAux(4).Text = Cad
        txtAux(3).Text = C2
    

End Select


End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function


    If FechaCorrecta2(CDate(txtAux(2).Text), True) > 1 Then Exit Function
    
    


    If Modo = 1 Then
        'Estamos insertando
         Datos = DevuelveDesdeBD("codigo", "stransfer", "codigo", txtAux(0).Text, "T")
         If Datos <> "" Then
            MsgBox "Ya existe la transferencia : " & txtAux(0).Text, vbExclamation
            B = False
        End If
    End If
    DatosOk = B
End Function


    


Private Function SepuedeBorrar() As Boolean
'Dim SQL As String
'    SepuedeBorrar = False
'    SQL = DevuelveDesdeBD("tipoamor", "samort", "numdiari", adodc1.Recordset!numdiari, "N")
'    If SQL <> "" Then
'        MsgBox "Esta vinculada con parametros de amortizacion", vbExclamation
'        Exit Function
'    End If
    
    SepuedeBorrar = True
End Function


Private Function GeneraNormaBancaria() As Boolean
Dim B As Boolean

    On Error GoTo EGeneraNormaBancaria
    GeneraNormaBancaria = False
    
    
    'Comprobamos las cuentas del banco de los recibos
    Set miRsAux = New ADODB.Recordset
    
    'Que el banco este bien
    If TipoDeFrm < 2 Then
        If Not comprobarCuentasBancariasPagos(CStr(adodc1.Recordset!Codigo), TipoDeFrm <> 0) Then
            Set miRsAux = Nothing
            Exit Function
        End If
    End If
    
    If Not ComprobarNifDatosProveedor Then Exit Function
        
   
    'Si es Norma34, transferenia, ofertaremos si queremos que el fichero sea
   
    Set miRsAux = Nothing
    If TipoDeFrm < 2 Then
        'B = GeneraFicheroNorma34(NIF, Adodc1.Recordset!Fecha, Adodc1.Recordset!codmacta, "9", Adodc1.Recordset!Codigo, Adodc1.Recordset!descripcion, TipoDeFrm <> 0)
        B = GeneraFicheroNorma34(NIF, adodc1.Recordset!Fecha, adodc1.Recordset!codmacta, CStr(adodc1.Recordset!conceptoTrans), adodc1.Recordset!Codigo, adodc1.Recordset!Descripcion, TipoDeFrm <> 0)
    
    Else
        
        
        If vParam.PagosConfirmingCaixa Then
            'Van por una "norma" de la caixa. De momento picassent
            B = GeneraFicheroCaixaConfirming(NIF, adodc1.Recordset!Fecha, adodc1.Recordset!codmacta, adodc1.Recordset!Codigo, adodc1.Recordset!Descripcion)
        Else
            'Q68
            'Fontenas, herbelca....
            B = GeneraFicheroNorma68(NIF, adodc1.Recordset!Fecha, adodc1.Recordset!codmacta, adodc1.Recordset!Codigo, adodc1.Recordset!Descripcion)
        End If
    End If
    If B Then
        If cd1.FileName <> "" Then cd1.FileName = ""
        cd1.ShowSave
        If cd1.FileName <> "" Then
            If Dir(cd1.FileName, vbArchive) <> "" Then
                If MsgBox("El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
                Kill cd1.FileName
            End If
        
            'CopiarFicheroNorma43 (TipoDeFrm < 2), cd1.FileName
            CopiarFicheroNormaBancaria TipoDeFrm, cd1.FileName
        End If
    Else
        MsgBox "Error generando fichero", vbExclamation
    End If
    
    
    Exit Function
EGeneraNormaBancaria:
    MuestraError Err.Number, "Genera Fichero Norma34"
End Function



Private Function ComprobarNifDatosProveedor() As Boolean
Dim SQL As String
        ComprobarNifDatosProveedor = False
        SQL = "select nifdatos,codmacta,nommacta from spagop,cuentas where spagop.ctaprove =cuentas.codmacta and  transfer = " & Me.adodc1.Recordset.Fields(0)
        SQL = SQL & " GROUP BY 1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            
            If Trim(DBLet(miRsAux!nifdatos, "T")) = "" Then SQL = SQL & "- " & miRsAux!codmacta & " " & miRsAux!Nommacta & vbCrLf
            miRsAux.MoveNext
        
        Wend
        miRsAux.Close
        If SQL <> "" Then
            MsgBox "Error en NIFs: " & vbCrLf & SQL, vbExclamation
            Set miRsAux = Nothing
        Else
            ComprobarNifDatosProveedor = True
        End If
        

End Function
