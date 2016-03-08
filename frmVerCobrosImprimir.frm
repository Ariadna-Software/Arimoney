VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerCobrosImprimir 
   Caption         =   "Imprimir recibos"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmVerCobrosImprimir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   5160
      Width           =   11175
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   9360
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL RIESGO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL PENDIENTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   11
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL VENCIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   3840
         TabIndex        =   9
         Top             =   120
         Width           =   1500
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerCobrosImprimir.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerCobrosImprimir.frx":5C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerCobrosImprimir.frx":5F48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7435
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
      NumItems        =   0
   End
   Begin VB.Frame frame 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   9480
         Picture         =   "frmVerCobrosImprimir.frx":6262
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprimir"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fec. VTO."
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fec. factura"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cta Cliente"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   8520
         Picture         =   "frmVerCobrosImprimir.frx":CAB4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Actualizar"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Cliente"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   8040
         Picture         =   "frmVerCobrosImprimir.frx":13306
         ToolTipText     =   "Seleccionar todos"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   7560
         Picture         =   "frmVerCobrosImprimir.frx":13450
         ToolTipText     =   "Quitar seleccion"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmVerCobrosImprimir.frx":1359A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Menu mnContextual 
      Caption         =   "Contextual"
      Visible         =   0   'False
      Begin VB.Menu mnNumero 
         Caption         =   "Poner numero Talón/Pagaré"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSelectAll 
         Caption         =   "Seleccionar todos"
      End
      Begin VB.Menu mnQUitarSel 
         Caption         =   "Quitar selección"
      End
   End
End
Attribute VB_Name = "frmVerCobrosImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vSQL As String
    'Diversas utilidades
    '-------------------------------------------------------------------------------
    'Para las transferencias me dice que transferencia esta siendo creada/modificada
    '
    'Para mostrar un check con los efectos k se van a generar en remesa y/o pagar

 
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Cad As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Fecha As Date
Dim Importe As Currency
Dim Vencido As Currency
Dim impo As Currency
Dim riesgo As Currency
Dim ImpSeleccionado  As Currency
Dim I As Integer
Private PrimeraVez As Boolean
Private SeVeRiesgo As Boolean
Dim SubItemVto











Private Sub cmdimprimir_Click()


    'Vamos a proceder a la impresion de los recibos
    
    Cad = ""
    For I = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Cad = Cad & "1"
    Next I
    If Cad = "" Then
        MsgBox "Deberias selecionar algun vencimiento.", vbExclamation
        Exit Sub
    End If
    
    
    
    'IMPRIMIMOS
    Screen.MousePointer = vbHourglass
    
    
    
            If GenerarRecibos2 Then
                'textoherecibido
                'DevuelveCadenaPorTipo True, Cad
                'If Cad = "" Then Cad = "He recibido de:"
                Cad = ""
                Cad = "textoherecibido= """ & Cad & """|"
                'Imprimimos
                CadenaDesdeOtroForm = DevuelveNombreInformeSCRYST(6, "Recibo")
                frmImprimir.Opcion = 8
                frmImprimir.NumeroParametros = 1
                frmImprimir.OtrosParametros = Cad
                frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                frmImprimir.Show vbModal
                
         
                
            End If
    
    
    
   
    Screen.MousePointer = vbDefault
End Sub


Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    CargaList
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        SeVeRiesgo = True
        Me.Refresh
        espera 0.1
        'Cargamos el LIST
        CargaList
        
            
            
        
        '----------------------------
        ' PRUEBAS
'        Debug.Print "--------------------------"
'        Debug.Print "VSQL:   " & vSQL
'        Debug.Print "Cobros:   " & Cobros
'        Debug.Print "Ordenar efecto:   " & OrdenarEfecto
'        Debug.Print "Regresar:   " & Regresar
'        Debug.Print "vtextos:   " & vTextos
'        Debug.Print "Tipo:   " & Tipo
'        Debug.Print "2ºparam:   " & SegundoParametro
'        Debug.Print "contab trans:   " & ContabTransfer
'        Stop
    End If
    Screen.MousePointer = vbDefault
End Sub
 

Private Sub DevuelveCadenaPorTipo(Impresion As Boolean, ByRef Cad1 As String)

    'FALTA###
    Dim Tipo As Integer
     
    Cad1 = ""
    Select Case Tipo
    Case 0
        If Impresion Then
            Cad1 = "He recibido mediante efectivo de"
        Else
            Cad1 = "[EFECTIVO]"
        End If
        
    Case 1
        Cad1 = "[TRANSFERENCIA]"
    Case 2
        If Impresion Then
            Cad1 = "He recibido mediante TALON de"
        Else
            Cad1 = "[TALON]"
        End If
    Case 3
        If Impresion Then
            Cad1 = "He recibido mediante PAGARE de"
        Else
            Cad1 = "[PAGARE]"
        End If
    
    Case 4
        Cad1 = "[RECIBO BANCARIO]"
    
    Case 5
        Cad1 = "[CONFIRMING]"
    
    Case 6
        If Impresion Then
            Cad1 = "He recibido mediante TARJETA DE CREDITO de"
        Else
            Cad1 = "[TARJETA CREDITO]"
        End If
    
    Case Else
        
        
        Stop
    End Select
End Sub

Private Sub Form_Load()


    PrimeraVez = True
    Limpiar Me
    Me.Icon = frmPpal.Icon
    CargaIconoListview Me.ListView1
    ListView1.Checkboxes = True
    Text1.Enabled = True



        
    Caption = "Cobros pendientes"
    Me.Option1(0).Caption = "Clientes"
    Me.Option1(3).Caption = "Nombre cliente"

    
    CargaGuardaOrdenacion True
 
    
    
    
    
    
    ListView1.SmallIcons = Me.ImageList1
    Text1.Text = Format(Now, "dd/mm/yyyy")
    Text1.Tag = "'" & Format(Now, FormatoFecha) & "'"
    CargaColumnas
    
    
    
    
End Sub

Private Sub Form_Resize()
Dim I As Integer
Dim h As Integer
    If Me.WindowState = 1 Then Exit Sub  'Minimizar
    If Me.Height < 2700 Then Me.Height = 2700
    If Me.Width < 2700 Then Me.Width = 2700
    
    'Situamos el frame y demas
    Me.frame.Width = Me.Width - 120
    Me.Frame1.Left = Me.Width - 120 - Me.Frame1.Width
    Me.Frame1.Top = Me.Height - Frame1.Height - 360

    
    Me.ListView1.Top = Me.frame.Height + 60
    Me.ListView1.Height = Me.Frame1.Top - Me.ListView1.Top - 60
    Me.ListView1.Width = Me.frame.Width
    
    'Las columnas
    h = ListView1.Tag
    ListView1.Tag = ListView1.Width - ListView1.Tag - 320 'Del margen
    For I = 1 To Me.ListView1.ColumnHeaders.Count
        If InStr(1, ListView1.ColumnHeaders(I).Tag, "%") Then
            Cad = (Val(ListView1.ColumnHeaders(I).Tag) * (Val(ListView1.Tag)) / 100)
        Else
            'Si no es de % es valor fijo
            Cad = Val(ListView1.ColumnHeaders(I).Tag)
        End If
        Me.ListView1.ColumnHeaders(I).Width = Val(Cad)
    Next I
    ListView1.Tag = h
End Sub


Private Sub CargaColumnas()
Dim colX As ColumnHeader
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim NCols As Integer
Dim I As Integer

    ListView1.ColumnHeaders.Clear
   
   
        NCols = 11
        Columnas = "Serie|Nº Factura|F. Fact|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
        Ancho = "700|13%|12%|12%|400|26%|660|11%|8%|11%|11%|"
        ALIGN = "LLLLLLLDDDD"
        
        
        ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
        
        
        'FALTA###
'        If Tipo = 2 Or Tipo = 3 Then
'            ''Si es un talon o pagare entonces añadire un campo mas
'            NCols = NCols + 1
'            Columnas = Columnas & "Nº Documento|"
'            Ancho = Ancho & "2500|"
'            ALIGN = ALIGN & "L"
'
'
'        End If
'

        
   For I = 1 To NCols
        Cad = RecuperaValor(Columnas, I)
        If Cad <> "" Then
            Set colX = ListView1.ColumnHeaders.Add()
            colX.Text = Cad
            'ANCHO
            Cad = RecuperaValor(Ancho, I)
            colX.Tag = Cad
            'align
            Cad = Mid(ALIGN, I, 1)
            If Cad = "L" Then
                'NADA. Es valor x defecto
            Else
                If Cad = "D" Then
                    colX.Alignment = lvwColumnRight
                Else
                    'CENTER
                    colX.Alignment = lvwColumnCenter
                End If
            End If
        End If
    Next I
End Sub


Private Sub CargaList()
On Error GoTo ECargando

    Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass

    SeVeRiesgo = True

    Label2(2).Visible = SeVeRiesgo
    Text2(2).Visible = SeVeRiesgo
    
    Set Rs = New ADODB.Recordset
    Fecha = CDate(Text1.Text)
    ListView1.ListItems.Clear
    Importe = 0
    Vencido = 0
    riesgo = 0
    ImpSeleccionado = 0
   
        CargaCobros

    
ECargando:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Text2(0).Text = Format(Vencido, FormatoImporte)
    Text2(1).Text = Format(Importe, FormatoImporte)
    Text2(2).Text = Format(riesgo, FormatoImporte)
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set Rs = Nothing
End Sub

Private Sub CargaCobros()
Dim Inserta As Boolean
Dim vImporte As Currency


    Cad = DevSQL
    
    'ORDENACION
    Cad = Cad & " ORDER BY "
    If Option1(0).Value Or Option1(3).Value Then
        'CLIENTE
        If Option1(0).Value Then
            'Codmacta
            Cad = Cad & " scobro.codmacta"
        Else
            Cad = Cad & " nommacta"
        End If
        Cad = Cad & ",numserie,codfaccl,fecfaccl"
    Else
        'FECHA FACTURA
        If Option1(1).Value Then
            Cad = Cad & " fecfaccl,numserie,codfaccl,fecvenci"
        Else
            Cad = Cad & " fecvenci,numserie,codfaccl,fecfaccl"
        End If
    End If
    'La ultima ordenacion por vto
    Cad = Cad & ",numorden"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Inserta = True
        If Rs!tipoformapago = vbTipoPagoRemesa Then
            
                    If DBLet(Rs!CodRem, "N") > 0 Then
                        Inserta = False
                        'Stop
                    End If
 
            
        
        End If
        
        If Inserta Then
    
    
            Set ItmX = ListView1.ListItems.Add()
            
            ItmX.Text = Rs!NUmSerie
            ItmX.SubItems(1) = Rs!codfaccl
            ItmX.SubItems(2) = Format(Rs!fecfaccl, "dd/mm/yyyy")
            ItmX.SubItems(3) = Format(Rs!fecvenci, "dd/mm/yyyy")
            ItmX.SubItems(4) = Rs!numorden
            ItmX.SubItems(5) = Rs!Nommacta
            ItmX.SubItems(6) = Rs!siglas
            
            ItmX.SubItems(7) = Format(Rs!impvenci, FormatoImporte)
            vImporte = DBLet(Rs!Gastos, "N")
            
            'Gastos
            ItmX.SubItems(8) = Format(vImporte, FormatoImporte)
            vImporte = vImporte + Rs!impvenci
            
            If Not IsNull(Rs!impcobro) Then
                ItmX.SubItems(9) = Format(Rs!impcobro, FormatoImporte)
                impo = vImporte - Rs!impcobro
                ItmX.SubItems(10) = Format(impo, FormatoImporte)
            Else
                impo = vImporte
                ItmX.SubItems(9) = "0.00"
                ItmX.SubItems(10) = Format(vImporte, FormatoImporte)
            End If
            If Rs!tipoformapago = vbTipoPagoRemesa Then
                '81--->
                'asc("Q") =81
                If Asc(Right(" " & DBLet(Rs!siturem, "T"), 1)) = 81 Then
                    riesgo = riesgo + vImporte
                Else
                   ' Stop
                End If
            End If
            If Rs!fecvenci < Fecha Then
                'LO DEBE
                ItmX.SmallIcon = 1
                Vencido = Vencido + impo
            Else
                ItmX.SmallIcon = 2
            End If
            Importe = Importe + impo
            
            ItmX.Tag = Rs!codmacta
            'FALTA###
            'If Tipo = 1 And SegundoParametro <> "" Then
            '    If Not IsNull(Rs!transfer) Then
            '        ItmX.Checked = True
            '        ImpSeleccionado = ImpSeleccionado + impo
            '    End If
            'End If
            
            
        End If  'de insertar
        Rs.MoveNext
    Wend
    Rs.Close
End Sub

Private Function DevSQL() As String
Dim Cad As String


        'cobros
        Cad = "SELECT scobro.*, sforpa.nomforpa, stipoformapago.descformapago, stipoformapago.siglas, "
        Cad = Cad & " cuentas.nommacta,cuentas.codmacta,stipoformapago.tipoformapago "
        Cad = Cad & " FROM ((scobro INNER JOIN sforpa ON scobro.codforpa = sforpa.codforpa) INNER JOIN stipoformapago ON sforpa.tipforpa = stipoformapago.tipoformapago) INNER JOIN cuentas ON scobro.codmacta = cuentas.codmacta"
        If vSQL <> "" Then Cad = Cad & " WHERE " & vSQL

    DevSQL = Cad
End Function



Private Sub Form_Unload(Cancel As Integer)
 
    CargaGuardaOrdenacion False
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1.Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub Image1_Click(Index As Integer)
    frmPreguntas.Opcion = 2
    frmPreguntas.Show vbModal
End Sub

Private Sub imgCheck_Click(Index As Integer)
    SeleccionarTodos Index = 1 Or Index = 2
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Fecha = Now
    If Index = 1 Then
        If Text1.Text <> "" Then
            If IsDate(Text1.Text) Then Fecha = CDate(Text1.Text)
        End If
    Else
'        If Text4.Text <> "" Then
'            If IsDate(Text4.Text) Then Fecha = CDate(Text4.Text)
'        End If
    End If
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Fecha
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then
        If Index = 1 Then
            Text1.Text = Cad
        Else
            'Text4.Text = Cad
        End If
    End If
End Sub



Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    I = ColD(0)
    impo = ImporteFormateado(Item.SubItems(I))
    
    If Item.Checked Then
        Set ListView1.SelectedItem = Item
        I = 1
    Else
        I = -1
    End If
    ImpSeleccionado = ImpSeleccionado + (I * impo)
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnContextual
    End If
End Sub

Private Sub SeleccionarTodos(Seleccionar As Boolean)
Dim J As Integer
    J = ColD(0)
    ImpSeleccionado = 0
    For I = 1 To Me.ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = Seleccionar
        impo = ImporteFormateado(ListView1.ListItems(I).SubItems(J))
        ImpSeleccionado = ImpSeleccionado + impo
    Next I
    If Not Seleccionar Then ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
End Sub


Private Sub mnNumero_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    CadenaDesdeOtroForm = "####"
    frmPreguntas.Opcion = 0
    frmPreguntas.vTexto = ListView1.SelectedItem.SubItems(11)
    frmPreguntas.Show vbModal
    If CadenaDesdeOtroForm <> "####" Then ListView1.SelectedItem.SubItems(11) = CadenaDesdeOtroForm
        
End Sub

Private Sub mnQUitarSel_Click()
    SeleccionarTodos False
End Sub

Private Sub mnSelectAll_Click()
    SeleccionarTodos True
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus()
    If Not EsFechaOK(Text1) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub


Private Function GenerarRecibos2() As Boolean
 Dim SQL As String
Dim Contador As Integer
Dim J As Integer
Dim Poblacion As String

    On Error GoTo EGenerarRecibos
    GenerarRecibos2 = False
    
    'Limpiamos
    Cad = "Delete from Usuarios.zTesoreriaComun where codusu = " & vUsu.Codigo
    Conn.Execute Cad


    'Guardamos datos empresa
    Cad = "Delete from Usuarios.z347carta where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos, "
    Cad = Cad & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, contacto, Asunto, Referencia)"
    Cad = Cad & " VALUES (" & vUsu.Codigo & ", "
    
    'Estos datos ya veremos com, y cuadno los relleno
    Set miRsAux = New ADODB.Recordset
    SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
    SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'"
    If Not miRsAux.EOF Then
        SQL = ""
        For J = 1 To 6
            SQL = SQL & DBLet(miRsAux.Fields(J), "T") & " "
        Next J
        SQL = Trim(SQL)
        SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
        Poblacion = DevNombreSQL(DBLet(miRsAux!Poblacion, "T"))
        SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & Poblacion & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Cad = Cad & SQL
    'otralinea,saludos
    Cad = Cad & ",NULL"
    'parrafo1
    Cad = Cad & ",''"
    
    
    '------------------------------------------------------------------------
    Cad = Cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
    Conn.Execute Cad

    'Empezamos     12 Mayo 2010. Añadimos el campo Texto de tesoreria comun donde pondremos todos razon social
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, "
    SQL = SQL & "texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion,texto)"
    SQL = SQL & " VALUES (" & vUsu.Codigo & ","

    'Julio 2009
    'Añadimos datos cliente desde vtos. scobro. nomclien,domclien,pobclien,cpclien,proclien
    ' IRan:   text5:  nomclien
    '         texto6: domclien
    '         observa2  cpclien  pobclien    + vbcrlf + proclien
    Set miRsAux = New ADODB.Recordset
    Contador = 0
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            'Lo insertamos tres veces
            
            RellenarCadenaSQLRecibo I, Poblacion
            

                Contador = Contador + 1
                Conn.Execute SQL & Contador & "," & Cad

         End If
    Next I
    Set miRsAux = Nothing
    GenerarRecibos2 = True
EGenerarRecibos:
    If Err.Number <> 0 Then
        MuestraError Err.Number
    End If
    Set miRsAux = Nothing
End Function


'----------------------------------
'Rellenaremos las cadenas de texto
Private Sub RellenarCadenaSQLRecibo(NumeroItem As Integer, Lugar As String)
Dim AUX As String
Dim QueDireccionMostrar As Byte
    '0. NO tiene
    '1. La del recibo
    '2. La de la cuenta

    
    With ListView1.ListItems(NumeroItem)
    
        ' IRan:   text5:  nomclien
        '         texto6: domclien
        '         observa2  cpclien  pobclien    + vbcrlf + proclien
    
        Cad = "select nomclien,domclien,pobclien,cpclien,proclien,razosoci,dirdatos,codposta,despobla,desprovi"
        'MAYO 2010
        Cad = Cad & ",codbanco,codsucur,digcontr,scobro.cuentaba,scobro.codmacta"
        'SEPTIEMBRE 2015
        Cad = Cad & ", text33csb ,  text41csb ,   scobro.obs"
        
        Cad = Cad & " from scobro,cuentas where scobro.codmacta =cuentas.codmacta and"
        Cad = Cad & " numserie ='" & .Text & "' and codfaccl=" & .SubItems(1)
        Cad = Cad & " and fecfaccl='" & Format(.SubItems(2), FormatoFecha) & "' and numorden=" & .SubItems(4)
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not miRsAux.EOF Then
            'El vto NO tiene datos de fiscales
            AUX = DBLet(miRsAux!nomclien, "T")
            If AUX = "" Then
                'La cuenta tampoco los tiene
                If IsNull(miRsAux!dirdatos) Then
                    QueDireccionMostrar = 0
                Else
                    QueDireccionMostrar = 2
                End If
            Else
                QueDireccionMostrar = 1
            End If
        Else
            QueDireccionMostrar = 0
        End If
        
        'texto1 , texto2, texto3, texto4, texto5,
        'texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        
            
        'Textos
        '---------
        '1.- Recibo nª
        Cad = "'" & .Text & "/" & .SubItems(1) & "'"
        
        'Pagos: cad = "'" & .Text & "/" & .SubItems(3) & "'"
        
        
        'Lugar Vencimiento
        Cad = Cad & ",'" & Lugar & "'"
        
        'text3 mostrare el codmacta
        'Cad = Cad & ",'" & DevNombreSQL(.SubItems(5)) & "',"
        Cad = Cad & ",'" & DevNombreSQL(miRsAux!codmacta) & "',"
        
        'MAYO 2010.    Ahora en este campo ira el CCC del cliente si es que lo tiene
        'Cad = Cad & "'" & .SubItems(6) & "'," ANTES
        AUX = DBLet(miRsAux!codbanco, "N")
        If AUX = "" Or AUX = "0" Then
            AUX = "NULL"
        Else
            'codbanco,codsucur,digcontr,cuentaba
            AUX = Format(DBLet(miRsAux!codbanco, "N"), "0000")
            AUX = AUX & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
            AUX = AUX & Mid(DBLet(miRsAux!digcontr, "T") & "  ", 1, 2) & " "
            AUX = AUX & Right(String(10, "0") & DBLet(miRsAux!cuentaba, "N"), 10)
            AUX = "'" & AUX & "'"
        End If
        Cad = Cad & AUX & ","
    
        '5 y 6.
        'text5: nomclien
        'texto6:domclien
        If QueDireccionMostrar = 0 Then
            'Cad = Cad & "NULL,NULL"
            'Siempre el nomclien
            Cad = Cad & "'" & DevNombreSQL(.SubItems(5)) & "',NULL"
        Else
            If QueDireccionMostrar = 1 Then
                Cad = Cad & "'" & DevNombreSQL(DBLet(miRsAux!nomclien, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!domclien, "T")) & "'"
            Else
                Cad = Cad & "'" & DevNombreSQL(DBLet(miRsAux!razosoci, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "'"
            End If
        End If
        
        Importe = ImporteFormateado(.SubItems(10))
        
        'IMPORTES
        '--------------------
        Cad = Cad & "," & TransformaComasPuntos(CStr(Importe))
        
        'El segundo importe NULL
        Cad = Cad & ",NULL"
        
        'FECFAS
        '--------------
        'Libramiento o pago     Auqi pone NOW
        'Ahora
        Cad = Cad & ",'" & Format(Text1.Text, FormatoFecha) & "'"
        'Antes de mayo 2013
        'Cad = Cad & ",'" & Format(Now, FormatoFecha) & "'"
        Cad = Cad & ",'" & Format(.SubItems(3), FormatoFecha) & "'"
        
        '3era fecha  NULL
        Cad = Cad & ",NULL"
        
        'OBSERVACIONES
        '------------------
        AUX = EscribeImporteLetra(Importe)
        
        AUX = "       ** " & AUX
        Cad = Cad & ",'" & AUX & "**',"
        
        
        'Observa 2
        '         observa2:    cpclien  pobclien    + vbcrlf + proclien
        If QueDireccionMostrar = 0 Then
            AUX = "NULL"
        Else
            
            If QueDireccionMostrar = 1 Then
                AUX = DBLet(miRsAux!cpclien, "T") & "      " & DevNombreSQL(DBLet(miRsAux!pobclien, "T"))
                AUX = Trim(AUX)
                If AUX <> "" Then AUX = AUX & vbCrLf
                AUX = AUX & DevNombreSQL(DBLet(miRsAux!proclien, "T"))
            Else
                AUX = DBLet(miRsAux!codposta, "T") & "      " & DevNombreSQL(DBLet(miRsAux!despobla, "T"))
                AUX = Trim(AUX)
                If AUX <> "" Then AUX = AUX & vbCrLf
                AUX = AUX & DevNombreSQL(DBLet(miRsAux!desprovi, "T"))
                
            End If
            AUX = "'" & AUX & "'"
        End If
        Cad = Cad & AUX
        
        
        
        'OPCION
        '--------------
        Cad = Cad & ",NULL,"
        
        
        'Septiembre 2015
        'En el campo observaciones del recibo (TEXT) guardaremos
        'text33csb(80)  text41csb(60)  scobro.obs(150)"
        AUX = Mid(DBLet(miRsAux!text33csb, "T") & Space(80), 1, 80)
        AUX = AUX & Mid(DBLet(miRsAux!text41csb, "T") & Space(60), 1, 60)
        AUX = AUX & Mid(DBLet(miRsAux!obs, "T") & Space(150), 1, 150)
        AUX = DevNombreSQL(AUX)
        Cad = Cad & " '" & AUX & "')"
        
        
    End With
    miRsAux.Close
End Sub
















'Es un jaleo. Cada vez que toque algo la vamos a liar
'Private Function ContabilizarLosPagos() As Boolean
'Dim J As Integer
'Dim Cuenta As String
'Dim MC2 As Contadores
'Dim GeneraAsiento As Boolean
'Dim Linea As Integer
'
'Dim UltimaAmpliacion As String
'Dim MismaCta As Boolean
'Dim ContraPartidaPorLinea As Boolean
'Dim UnAsientoPorCuenta As Boolean
'Dim LineasCuenta As Integer
'Dim vGasto As Currency
'
'Dim AgrupaCuentaGenerica As Boolean
'Dim CtaAgrupada As String
'
'
'Dim OtraCuenta As Boolean
'
'    On Error GoTo ECon
'
'    FechaAsiento = CDate(Text3(0).Text)
'
'    ContraPartidaPorLinea = True
'    UnAsientoPorCuenta = False
'    AgrupaCuentaGenerica = False
'    If Tipo = 1 And ContabTransfer Then
'        If Me.chkContrapar(1).Value Then ContraPartidaPorLinea = False
'        If Me.chkAsiento(1).Value Then UnAsientoPorCuenta = True
'        If chkGenerico(1).Value Then AgrupaCuentaGenerica = True
'        CtaAgrupada = chkGenerico(1).Tag
'    Else
'        'Si no es transferencia
'        If Me.chkContrapar(0).Value Then ContraPartidaPorLinea = False
'        If Me.chkAsiento(0).Value Then UnAsientoPorCuenta = True
'        If chkGenerico(0).Value Then AgrupaCuentaGenerica = True
'        CtaAgrupada = chkGenerico(0).Tag
'    End If
'
'
'
'
'    ContabilizarLosPagos = False
'    Cuenta = ""
'    Set MC2 = New Contadores
'
'    Stop   'Paramos para ver esto bien
'
'    MismaCta = True
'    vGasto = 0
'    riesgo = 0
'    OtraCuenta = False
'
'    For J = 1 To ListView1.ListItems.Count
'       If ListView1.ListItems(J).Checked Then
'
'            'Veremos si es otra cuenta o no
'            If AgrupaCuentaGenerica Then
'                If Cuenta = "" Then
'                    OtraCuenta = True 'Para que genere el asiento
'                Else
'                    OtraCuenta = False
'                End If
'            Else
'                If ListView1.ListItems(J).Tag <> Cuenta Then OtraCuenta = True
'            End If
'
'            'If ListView1.ListItems(J).Tag <> Cuenta Then
'            If OtraCuenta Then
'                If Cuenta = "" Then
'                    GeneraAsiento = True
'                Else
'                    'SI en PARAMETROS pone k hay nuevo asiento por pago
'                    'Entonces
'
'                    If UnAsientoPorCuenta Then
'                        GeneraAsiento = True
'                    Else
'                        GeneraAsiento = False
'                    End If
'                End If
'
'
'                'Saldamos la cuenta de banco con respecto al cliente
'                '-------------------------------------------------------
'                '-------------------------------------------------------
'                ' Con respecto al cliente. Es decir:
'                '   - Si estoy cerrando por contrapartida NO tendre que      esto
'                If UnAsientoPorCuenta Then
'                    If Cuenta <> "" Then
'                        impo = Importe  'Para el importe
'                        'No va la J, k sera del nuevo cli/pro
'                        'Si no k va J-1
'
'                        Linea = Linea + 1
'                        If Not ContraPartidaPorLinea Then
'                            'Hay mas de una de banco, con lo cual, NO hace referencia a nada el documento
'                            If LineasCuenta > 1 Then UltimaAmpliacion = ""
'                            'Insertamos el banco
'                            InsertarEnAsientos MC2, Linea, J - 1, 2, Cuenta, UltimaAmpliacion, LineasCuenta = 1, AgrupaCuentaGenerica, CtaAgrupada
'                        End If
'                        'Lo insertamos en tmpactualizar
'                        If GeneraAsiento Then InsertarEnAsientos MC2, Linea, J - 1, 3, Cuenta, UltimaAmpliacion, False, False, CtaAgrupada
'                        riesgo = 0
'                        Importe = 0
'                        vGasto = 0
'                    End If
'                End If
'                If GeneraAsiento Then
'
'                    UltimaAmpliacion = ""   'Por si salda uno a uno los pagos
'                    'ES el primero.
'                    'Obtener contador
'                    If MC2.ConseguirContador("0", FechaAsiento <= vParam.fechafin, True) = 1 Then Exit Function
'                     'Es la cabecera. La primera I no la tratamos en cabecera
'                    InsertarEnAsientos MC2, I, I, 0, Cuenta, UltimaAmpliacion, False, False, CtaAgrupada
'                    Importe = 0
'                    Linea = 0
'                    riesgo = 0
'                    vGasto = 0
'                    MismaCta = True
'                End If
'                'If Cuenta <> "" Then MismaCta = (Cuenta = ListView1.ListItems(J).Tag)
'                Cuenta = ListView1.ListItems(J).Tag
'                LineasCuenta = 0
'            End If
'            I = ColD(0)
'            impo = ImporteFormateado(ListView1.ListItems(J).SubItems(I))
'
'            'riesgo es GASTO
'            If Cobros Then
'                riesgo = ImporteFormateado(ListView1.ListItems(J).SubItems(I - 2))
'            Else
'                riesgo = 0
'            End If
'            impo = impo - riesgo
'            vGasto = vGasto + riesgo
'
'            Importe = Importe + impo
'
'            Linea = Linea + 1
'
'            InsertarEnAsientos MC2, Linea, J, 1, Cuenta, UltimaAmpliacion, False, AgrupaCuentaGenerica, CtaAgrupada
'
'
'            LineasCuenta = LineasCuenta + 1
'            'Si es cobros o pagos
'            If ContraPartidaPorLinea Then
'                Linea = Linea + 1
'                InsertarEnAsientos MC2, Linea, J, 2, Cuenta, UltimaAmpliacion, True, AgrupaCuentaGenerica, CtaAgrupada
'                Importe = 0
'            End If
'
'        End If
'    Next J
'    'Nos faltara cerrar la ultima linea de banco caja
'    impo = Importe  'Para el importe
'
'    'No va la J, k sera del nuevo cli/pro
'    'Si no k va J-1
'    Linea = Linea + 1
'
'    If ContraPartidaPorLinea Then
'        If MismaCta Then
'            'If LineasCuenta > 1 Then UltimaAmpliacion = Cuenta
'            If LineasCuenta > 1 Then UltimaAmpliacion = ""
'        End If
'    Else
'        'Si creo solo un apunte banco por asiento, no pongo ampliacion ni doumento
'        UltimaAmpliacion = ""
'    End If
'    riesgo = vGasto
'    If impo <> 0 Then InsertarEnAsientos MC2, Linea, J - 1, 2, Cuenta, UltimaAmpliacion, LineasCuenta = 1, AgrupaCuentaGenerica, CtaAgrupada  'Genera
'
'    InsertarEnAsientos MC2, Linea, J - 1, 3, Cuenta, UltimaAmpliacion, False, False, CtaAgrupada 'Cerramos el asiento
'
'
'
'    'Todo OK
'    ContabilizarLosPagos = True
'
'
'ECon:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, Err.Description
'    End If
'    Set MC2 = Nothing
'
'End Function



' A partir de un numero de columna nos dira k columna es
' en el LISTVIEW
'
Private Function ColD(Colu As Integer) As Integer
    Select Case Colu
    Case 0
            'IMporte pendiente
            ColD = 10
    Case 1
    
    End Select
    
End Function















Private Sub CargaGuardaOrdenacion(Leer As Boolean)

    On Error GoTo ECargaGuardaOrdenacion
'
'
'    Cad = App.Path & "\ordeefec.xdf"
'    I = FreeFile
'    If Leer Then
'        OrdenacionEfectos = 0
'        If Dir(Cad, vbArchive) <> "" Then
'            Open Cad For Input As #I
'            Line Input #I, Cad
'            Close #I
'            If Cad <> "" Then
'                I = Val(Cad)
'                If I > 3 Then I = 0
'                OrdenacionEfectos = I
'            End If
'        End If
'
'
'    Else
'        'guardar
'        SubItemVto = 0
'        For I = 0 To 3
'            If Me.Option1(I).Value Then SubItemVto = I
'        Next I
'
'        If SubItemVto <> OrdenacionEfectos Then
'
'
'            If SubItemVto = 0 Then
'                If Dir(Cad, vbArchive) <> "" Then Kill Cad
'            Else
'                Open Cad For Output As #I
'                Print #I, SubItemVto
'                Close #I
'            End If
'        End If
'
'
'    End If
'    Exit Sub
    
ECargaGuardaOrdenacion:
    Err.Clear
End Sub








