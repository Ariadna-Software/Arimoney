VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión listados"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
        '0 .- Listado cobros pendientes por CLIENTE
    
    
    
        '7.- Reclamacion por CARTA
    
        '8.- Recibos
    
    
        '9.- Agentes
        '10.- Departamentos
        
        '11.- Remesas
    
        '12.- Listado de caja
    
        '13.- Impagados, por fecha
        '14.-   "        por cliente
    
    
        'Listado de cobros pendientes por cliente.
        '  Reservo hasta el
        '  20 y algo
    
    
        '25: Deuda agrupada por nif
    
        '26: Listado  bancos
        '26: Formas de pago
        '27
        
        '29: PREVISION
    
    
        'Operaciones aseguradas
        '
        '31:    Datos basicos
        '32:    Listado facturacion
        '33:    Impagados
        '34:    Listado efectos asegurados
        
        
        
        
        '40:    Carta confirming(TEINSA)
        '41:    Caja con SALDO arrastrado
        
        
        '50   : los reservo tb hasta el 55
        
        
        '55:  Hasta este reservado
        
        
        
        '60: Recibo pago proveedores
        
        '61: Listado de documentos recibidos (talones / pagares)  desglosado
        
        '62: Orden pago bancos
        
        '63.  docs recibidos SIN desglosar
        
        
        
        'Salto
        '*******************************
        
        '70 as 85 (me reservo unos pocos)
        '   Cobros pendientes que en lugar de por codmacta ira por nommacta
        '   71
        '......
        
        
        '86     Listado reclamaciones
        
        '87   Carta confirmacion recepcion documentos
        
        '88     Listado remesas talon pagare para llevarlas al banco
        
        '89     Listado gastos fijos
        
        '90     Operaciones aseguradas. listado avisos
        
        '91     Compensacion de clientes
        
        '92     Cominicacion datos seguro
        
        '93,94  Facturas pendientes Operaciones aseguradas
        
        'libres
        '...
        
                
        
        
        'Los 500 estan reservados para los informes APAISADOS de cobros
        '
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public EnvioEMail As Boolean


Public QueEmpresaEs As Byte
    '0 Todas
    '1 Escalon

Private MostrarTree As Boolean
Private Nombre As String
Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean


'Private ReestableceSoloImprimir As Boolean

Private Sub chkEmail_Click()
    If chkEmail.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEmail.Value = 0
End Sub

Private Sub cmdConfigImpre_Click()
Screen.MousePointer = vbHourglass
'Me.CommonDialog1.Flags = cdlPDPageNums
CommonDialog1.ShowPrinter
PonerNombreImpresora
Screen.MousePointer = vbDefault
End Sub


Private Sub cmdimprimir_Click()
If Me.chkSoloImprimir.Value = 1 And Me.chkEmail.Value = 1 Then
    MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation
    Exit Sub
End If
'Form2.Show vbModal
Imprime
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub Form_Activate()
If PrimeraVez Then
    espera 0.1
    CommitConexion
    If SoloImprimir Then
        Imprime
        Unload Me
    Else
        If EnvioEMail Then
            Me.Hide
            chkEmail.Value = 1
            Imprime
            Unload Me
        End If
    End If
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Cad As String
Dim InformesApaisados As Boolean   'De momento solo vale para los cobros
PrimeraVez = True
Lanzado = False
CargaICO
Cad = Dir(App.Path & "\impre.dat", vbArchive)


'ReestableceSoloImprimir = False
If Cad = "" Then
    chkSoloImprimir.Value = 0
    Else
    chkSoloImprimir.Value = 1
    'ReestableceSoloImprimir = True
End If
cmdImprimir.Enabled = True
If SoloImprimir Then
    chkSoloImprimir.Value = 0
    Me.Frame2.Enabled = False
    chkSoloImprimir.Visible = False
Else
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
PonerNombreImpresora
MostrarTree = False

'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)


    MIPATH = App.Path & "\InformesT\"
InformesApaisados = False
If Opcion > 500 Then
    InformesApaisados = True
    'FALTA
    'Iremos haciendo para cada informe de cobro el _A. Ejem: cobrospen.rpt  cobrospen_A.rpt
    ' y asi, haremos opcion-500  y se quedara la opcion k corresponde

End If
Select Case Opcion
Case 1, 71
    Text1.Text = "Cobros pendiente CLIENTE"
    Nombre = "cobrospenCLI"
    If Opcion = 71 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    Nombre = Nombre & ".rpt"
    MostrarTree = True
    
Case 2, 72
    Text1.Text = "Cobros pendiente CLIENTE por F. Vto"
    Nombre = "cobrospenCLIF"
    If Opcion = 72 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    Nombre = Nombre & ".rpt"

    MostrarTree = True
    
Case 3, 73
    Text1.Text = "Resumen cobros pendiente CLIENTE"
    Nombre = "cobrospenCLISIN"
    If Opcion = 73 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    Nombre = Nombre & ".rpt"

    
    
Case 4
    Text1.Text = "Pagos pendientes proveedores"
    Nombre = "pagospenPRO.rpt"
    
Case 5
    Text1.Text = "Pagos pendientes F. Vencimiento"
    Nombre = "pagospenPROF.rpt"
    
Case 6
    Text1.Text = "Resumen Pagos pendientes proveedor"
    Nombre = "pagospenPROSIN.rpt"

Case 7
    ' RECLAMACIONES
    'Es especial.
    'Si en CadenaDesdeOtroForm vien algo entonces ese será el INFORME
    'Si no... reclamas.rpt
    

    Text1.Text = "Reclamaciones cobros"
    If CadenaDesdeOtroForm <> "" Then
        Text1.Text = Text1.Text & "(" & CadenaDesdeOtroForm & ")"
        Nombre = CadenaDesdeOtroForm
        
    Else
        Nombre = "reclamas.rpt"
    End If
    
    
    

Case 8
    Text1.Text = "Recibos"
    If CadenaDesdeOtroForm = "" Then
        Nombre = "recibos1.rpt"
    Else
        Nombre = CadenaDesdeOtroForm
    End If
Case 9
    Text1.Text = "Agentes"
    Nombre = "Agentes.rpt"
Case 10
    Text1.Text = "Departamentos"
    Nombre = "Dptos.rpt"
    MostrarTree = True

Case 11
    Text1.Text = "Remesas"
    Nombre = "Remesas1.rpt"
Case 12
    Text1.Text = "Caja"
    Nombre = "Caja.rpt"
Case 13, 14
    Text1.Text = "Efectos devueltos por "
    If Opcion = 13 Then
        Nombre = 1
        Text1.Text = Text1.Text & " fecha"
    Else
        Nombre = 2
        Text1.Text = Text1.Text & " cliente"
    End If
    Nombre = "impaga" & Nombre & ".rpt"
    MostrarTree = True
    
    
Case 21, 74
    Text1.Text = "Cobros pendiente CLIENTE. Situación"
    Nombre = "cobrospenCLIsi"
    If Opcion = 74 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    Nombre = Nombre & ".rpt"
    MostrarTree = True
    
Case 22, 75
    Text1.Text = "Cobros pendiente CLIENTE por F. Vto. Situación"
    Nombre = "cobrospenCLIFsi"
    If Opcion = 75 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    Nombre = Nombre & ".rpt"
    MostrarTree = True
    
Case 23, 76

    Text1.Text = "Resumen cobros pendiente CLIENTE. Situación"
    Nombre = "cobrospenCLISINsi"
    If Opcion = 76 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & "(Nom)"
    End If
    Nombre = Nombre & ".rpt"

Case 25
    Text1.Text = "Cobros / Pagos agrupados por nif"
    Nombre = "niffecha.rpt"
    MostrarTree = True

Case 26
    Text1.Text = "Listado bancos"
    Nombre = "bancosprop.rpt"
    

Case 27
    Text1.Text = "Formas de pago"
    Nombre = "forpa.rpt"
    
Case 28
    Text1.Text = "Listado (tran/conf)"
    Nombre = "transfer1.rpt"
    
Case 29
    Text1.Text = "Previsión"
    Nombre = "Prevision1.rpt"
    MostrarTree = True

Case 31, 32, 33, 34
    
    If Opcion = 31 Then
        CadenaDesdeOtroForm = "Datos básicos"
        Nombre = "asegbasic.rpt"  '32
    Else
        If Opcion = 32 Then
            CadenaDesdeOtroForm = "Facturación"
            Nombre = "aseglistfac.rpt"
            MostrarTree = True
        Else
            If Opcion = 33 Then
                CadenaDesdeOtroForm = "Impagados"
                Nombre = "asegimpag.rpt"
            Else
                CadenaDesdeOtroForm = "Listado efectos"
                Nombre = "asegefect.rpt"
            End If
        End If
    End If
    Text1.Text = "Operaciones aseguradas. (" & CadenaDesdeOtroForm & ")"
    CadenaDesdeOtroForm = ""
Case 32
    
    MostrarTree = True
    
    
Case 40
    Text1.Text = "Pagaré (" & CadenaDesdeOtroForm & ")"
    Nombre = CadenaDesdeOtroForm
        
        
Case 41
    Text1.Text = "Caja con saldo arrastrado"
    Nombre = "cajaS.rpt"
        
        
        
'salto


Case 50 To 55, 77, 78, 79
    'Listado cobros pendeintes agrupados por forma pago
    Text1.Text = "Cobros pend. x Forma pago  / "
    MostrarTree = True
    If Opcion = 51 Or Opcion = 77 Then
        Nombre = "cobrospen_cli_fp"
        Text1.Text = Text1.Text & "cliente (desglose "
    ElseIf Opcion = 52 Or Opcion = 78 Then
        MostrarTree = False
        Nombre = "cobrospenCLIF_fp"
        Text1.Text = Text1.Text & " (Fecha"
    Else
        '53-79
        Nombre = "cobrospenCLISIN_fp"
        Text1.Text = Text1.Text & " (cliente"
    End If
    If Opcion >= 77 Then
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & " Nomb."
    End If
    Text1.Text = Text1.Text & ")"
    Nombre = Nombre & ".rpt"
    
    
Case 60
    Nombre = CadenaDesdeOtroForm
    Text1.Text = "Recibo pago proveedores"
    
Case 61, 63  'con o sin desglose
    Text1.Text = "Documentos recibidos"
    
    Nombre = "talonpag"
    If Opcion = 63 Then
        Nombre = Nombre & "2"
    Else
        Text1.Text = Text1.Text & " (Desglose"
        MostrarTree = True
    End If
    Nombre = Nombre & ".rpt"

Case 62
    
    Text1.Text = "Orden de pago banco"
    MostrarTree = True
    If CadenaDesdeOtroForm = "" Then
        Nombre = "rOrdenPago.rpt"
    Else
        Nombre = CadenaDesdeOtroForm
    End If

Case 86
    Text1.Text = "Listado reclamaciones"
    Nombre = "hcoReclamas.rpt"

Case 87
    Text1.Text = "Confirmación recepción documentos"
    Nombre = "talonpagConfRec.rpt"
     If CadenaDesdeOtroForm <> "" Then
        If UCase(Nombre) <> UCase(CadenaDesdeOtroForm) Then Text1.Text = Text1.Text & "(" & CadenaDesdeOtroForm & ")"
        Nombre = CadenaDesdeOtroForm
    Else
        Nombre = "talonpagConfRec.rpt"
    End If
Case 88
    Text1.Text = "Remesas T/P a banco"
    Nombre = "remesasBanco.rpt"
    
Case 89
    Text1.Text = "Gastos fijos"
    Nombre = "GastosFijos.rpt"
    
    
Case 90
    Text1.Text = "Listados de avisos aseguradoras"
    Nombre = "aseAvisos.rpt"
    
    
Case 91
    Text1.Text = "Listados de compensacion de cobros"
    Nombre = CadenaDesdeOtroForm
    
Case 92
    Text1.Text = "Comunicación aseguradora"
    'Nombre = "rComunicaSeguro.rpt"
    Nombre = CadenaDesdeOtroForm
    
Case 93, 94
    Text1.Text = "Fras. pendientes operaciones aseg."
    If Opcion = 94 Then
        Text1.Text = Text1.Text & "(RES)"
        Nombre = "rSeguroFrasPdtesRes.rpt"
    Else
        Nombre = "rSeguroFrasPdtes.rpt"
    End If
    
    
Case 95
    Text1.Text = "Carta abonos"
    'Nombre = "rCartatransfer.rpt"
    Nombre = CadenaDesdeOtroForm
    
Case 96
    Text1.Text = "Cobros agentes (Bacchus)"
    'Nombre = "rCartatransfer.rpt"
    Nombre = CadenaDesdeOtroForm
    
    
Case 500 To 599
    Text1.Text = "Apaisado"
    Nombre = "cobrospencli_A.rpt"
Case Else
    Text1.Text = "Opcion incorrecta"
    Me.cmdImprimir.Enabled = False
End Select



Screen.MousePointer = vbDefault
End Sub




Private Function Imprime() As Boolean
Dim Seguir As Boolean

    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & Nombre
        .ExportarPDF = (chkEmail.Value = 1)
        .Show vbModal
    End With
    
    If Me.chkEmail.Value = 1 Then
    
        'Para la opcion 7 (reclamaciones) NO envia desde aqui
        If Opcion <> 7 Then
            If CadenaDesdeOtroForm <> "" Then
                 frmEMail.queEmpresa = QueEmpresaEs
                 frmEMail.Opcion = 0
                 frmEMail.Show vbModal
            End If
            CadenaDesdeOtroForm = ""
        End If
    End If
    Unload Me
 
 
 
End Function


Private Sub Form_Unload(Cancel As Integer)
    If Me.chkEmail.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    OperacionesArchivoDefecto
    QueEmpresaEs = 0
    EnvioEMail = False
End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

crear = (Me.chkSoloImprimir.Value = 1)
'crear = crear And ReestableceSoloImprimir
If Not crear Then
    Kill App.Path & "\impre.dat"
    Else
        FileCopy App.Path & "\Vacio.dat", App.Path & "\impre.dat"
End If
ErrOperacionesArchivoDefecto:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Text1_DblClick()
Frame2.Tag = Val(Frame2.Tag) + 1
If Val(Frame2.Tag) > 2 Then
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    Label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub

Private Sub CargaICO()
    On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    If Err.Number <> 0 Then Err.Clear
End Sub

