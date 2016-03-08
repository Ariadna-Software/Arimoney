VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOrdenarPagos 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   8895
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
         Left            =   6720
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   0
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
         Index           =   0
         Left            =   2040
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL PENDIENTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL VENCIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   2535
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
            Picture         =   "frmOrdenarPagos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenarPagos.frx":5C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenarPagos.frx":5F3C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6800
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
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.CommandButton cmdRegresar 
         Height          =   615
         Left            =   8040
         Picture         =   "frmOrdenarPagos.frx":6256
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Regresar"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fec. VTO."
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fec. factura"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   1560
         Picture         =   "frmOrdenarPagos.frx":6C58
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
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmOrdenarPagos.frx":6F62
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
End
Attribute VB_Name = "frmOrdenarPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vSQL As String
Public Cobros As Boolean
Public OrdenarEfecto As Boolean
Public Regresar As Boolean
    'Para mostrar un check con los efectos k se van a generar en remesa y/o pagar

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Cad As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim Fecha As Date
Dim Importe As Currency
Dim Vencido As Currency
Dim Impo As Currency
Private PrimeraVez As Boolean

Private Sub cmdRegresar_Click()
    If Not (ListView1.SelectedItem Is Nothing) Then
        If Cobros Then
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.SelectedItem.SubItems(2) & "|" & ListView1.SelectedItem.SubItems(4) & "|"
        Else
            'Pagos proveedores
            CadenaDesdeOtroForm = ListView1.SelectedItem.Tag & "|" & ListView1.SelectedItem.Text & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.SelectedItem.SubItems(1) & "|" & ListView1.SelectedItem.SubItems(3) & "|"
        End If
    Else
        CadenaDesdeOtroForm = ""
    End If
    Unload Me
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    CargaList
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        'Cargamos el LIST
        CargaList
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Limpiar Me
    ListView1.Checkboxes = OrdenarEfecto
    If Cobros Then
        Caption = "Cobros pendientes"
        Me.Option1(0).Caption = "Clientes"
    Else
        Caption = "Pagos pendientes"
        Me.Option1(0).Caption = "Proveedores"
    End If
    Me.cmdRegresar.Visible = Regresar
    ListView1.SmallIcons = Me.ImageList1
    Text1.Text = Format(Now, "dd/mm/yyyy")
    Text1.Tag = "'" & Format(Now, FormatoFecha) & "'"
    CargaColumnas
    
End Sub

Private Sub Form_Resize()
Dim I As Integer
Dim H As Integer
    If Me.WindowState = 1 Then Exit Sub  'Minimizar
    If Me.Height < 2000 Then Me.Height = 2000
    If Me.Width < 2000 Then Me.Width = 2000
    
    'Situamos el frame y demas
    Me.frame.Width = Me.Width - 120
    Me.Frame1.Left = Me.Width - 120 - Me.Frame1.Width
    Me.Frame1.Top = Me.Height - Frame1.Height - 360
    Me.ListView1.Top = Me.frame.Height + 60
    Me.ListView1.Height = Me.Frame1.Top - Me.ListView1.Top - 60
    Me.ListView1.Width = Me.frame.Width
    
    'Las columnas
    H = ListView1.Tag
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
    ListView1.Tag = H
End Sub


Private Sub CargaColumnas()
Dim colX As ColumnHeader
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim NCols As Integer
Dim I As Integer

    ListView1.ColumnHeaders.Clear
   If Cobros Then
        NCols = 10
        Columnas = "Serie|Nº Factura|F. Fact|F. VTO|Nº|CLIENTE|Tipo|Importe|Cobrado|Pendiente|"
        Ancho = "800|13%|12%|12%|400|26%|700|12%|12%|12%|"
        ALIGN = "LLLLLLLDDD"
        ListView1.Tag = 2100  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
   Else
        NCols = 9
        Columnas = "Nº Factura|F. Fact|F. VTO|Nº|PROVEEDOR|Tipo|Importe|Cobrado|Pendiente|"
        Ancho = "15%|12%|12%|400|26%|700|12%|12%|12%|"
        ALIGN = "LLLLLLDDD"
        ListView1.Tag = 1500  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
    End If
        
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
    Set RS = New ADODB.Recordset
    Fecha = CDate(Text1.Text)
    ListView1.ListItems.Clear
    Importe = 0
    Vencido = 0
    If Cobros Then
        CargaCobros
    Else
        CargaPagos
    End If
    
    
ECargando:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Text2(0).Text = Format(Vencido, FormatoImporte)
    Text2(1).Text = Format(Importe, FormatoImporte)
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set RS = Nothing
End Sub

Private Sub CargaCobros()

    
    Cad = "SELECT scobro.*, sforpa.nomforpa, stipoformapago.descformapago, stipoformapago.siglas, "
    Cad = Cad & " cuentas.nommacta"
    Cad = Cad & " FROM ((scobro INNER JOIN sforpa ON scobro.codforpa = sforpa.codforpa) INNER JOIN stipoformapago ON sforpa.tipforpa = stipoformapago.tipoformapago) INNER JOIN cuentas ON scobro.codmacta = cuentas.codmacta"
    
    
    
    If vSQL <> "" Then Cad = Cad & " WHERE " & vSQL
    
    'ORDENACION
    Cad = Cad & " ORDER BY "
    If Option1(0).Value Then
        'CLIENTE
        Cad = Cad & " scobro.codmacta,numserie,codfaccl,fecfaccl"
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
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add()
        
        ItmX.Text = RS!numserie
        ItmX.SubItems(1) = RS!codfaccl
        ItmX.SubItems(2) = Format(RS!fecfaccl, "dd/mm/yyyy")
        ItmX.SubItems(3) = Format(RS!fecvenci, "dd/mm/yyyy")
        ItmX.SubItems(4) = RS!numorden
        ItmX.SubItems(5) = RS!nommacta
        ItmX.SubItems(6) = RS!siglas
        ItmX.SubItems(7) = Format(RS!impvenci, FormatoImporte)
        If Not IsNull(RS!impcobro) Then
            ItmX.SubItems(8) = Format(RS!impcobro, FormatoImporte)
            Impo = RS!impvenci - RS!impcobro
            ItmX.SubItems(9) = Format(Impo, FormatoImporte)
        Else
            Impo = RS!impvenci
            ItmX.SubItems(8) = "0.00"
            ItmX.SubItems(9) = ItmX.SubItems(7)
        End If
        If RS!fecvenci < Fecha Then
            'LO DEBE
            ItmX.SmallIcon = 1
            Vencido = Vencido + Impo
        Else
            ItmX.SmallIcon = 2
        End If
        Importe = Importe + Impo
        RS.MoveNext
    Wend
    RS.Close
End Sub


Private Sub CargaPagos()
    'Cad = "SELECT scobro.*, sforpa.nomforpa, stipoformapago.descformapago, stipoformapago.siglas, "
    'Cad = Cad & " cuentas.nommacta"
    'Cad = Cad & " FROM ((scobro INNER JOIN sforpa ON scobro.codforpa = sforpa.codforpa) INNER JOIN stipoformapago ON sforpa.tipforpa = stipoformapago.tipoformapago) INNER JOIN cuentas ON scobro.codmacta = cuentas.codmacta"
    
    
    Cad = "SELECT spagop.*, cuentas.nommacta, stipoformapago.siglas FROM"
    Cad = Cad & " spagop , cuentas, sforpa, stipoformapago"
    Cad = Cad & " Where spagop.ctaprove = cuentas.codmacta"
    Cad = Cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"
    Cad = Cad & " AND spagop.codforpa = sforpa.codforpa"
    
    
    If vSQL <> "" Then Cad = Cad & " AND " & vSQL
    
    'ORDENACION
    Cad = Cad & " ORDER BY "
    If Option1(0).Value Then
        'CLIENTE
        Cad = Cad & " spagop.ctaprove,numfactu,fecfactu"
    Else
        'FECHA FACTURA
        If Option1(1).Value Then
            Cad = Cad & " fecfactu,numfactu,fecefect"
        Else
            Cad = Cad & " fecefect,numfactu,fecfactu"
        End If
    End If
    'La ultima ordenacion por vto
    Cad = Cad & ",numorden"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add()
        
        ItmX.Text = RS!numfactu
        ItmX.SubItems(1) = Format(RS!fecfactu, "dd/mm/yyyy")
        ItmX.SubItems(2) = Format(RS!fecefect, "dd/mm/yyyy")
        ItmX.SubItems(3) = RS!numorden
        ItmX.SubItems(4) = RS!nommacta
        ItmX.SubItems(5) = RS!siglas
        ItmX.SubItems(6) = Format(RS!impefect, FormatoImporte)
        If Not IsNull(RS!imppagad) Then
            ItmX.SubItems(7) = Format(RS!imppagad, FormatoImporte)
            Impo = RS!impefect - RS!imppagad
            ItmX.SubItems(8) = Format(Impo, FormatoImporte)
        Else
            Impo = RS!impefect
            ItmX.SubItems(7) = "0.00"
            ItmX.SubItems(8) = ItmX.SubItems(6)
        End If
        If RS!fecefect < Fecha Then
            'LO DEBE
            ItmX.SmallIcon = 1
            Vencido = Vencido + Impo
        Else
            ItmX.SmallIcon = 2
        End If
        'El tag lo utilizo para la cta proveedor
        ItmX.Tag = RS!ctaprove
        
        Importe = Importe + Impo
        RS.MoveNext
    Wend
    RS.Close

End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Fecha = Now
    If Text1.Text <> "" Then
        If IsDate(Text1.Text) Then Fecha = CDate(Text1.Text)
    End If
    Set frmC = New frmCal
    frmC.Fecha = Fecha
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub ListView1_DblClick()
    If Not (ListView1.SelectedItem Is Nothing) Then
        If Regresar Then cmdRegresar_Click
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
    If Not EsFechaOK(Text1) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

