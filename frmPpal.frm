VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPpal 
   BackColor       =   &H00848684&
   Caption         =   "MDIForm1"
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12225
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmPpal.frx":42B2
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImgListviews 
      Left            =   4800
      Top             =   600
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
            Picture         =   "frmPpal.frx":24CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2B54E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2DD00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcoForms 
      Left            =   6840
      Top             =   480
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
            Picture         =   "frmPpal.frx":33922
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":34334
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":343CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar11 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tipos de pago"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Formas de pago"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bancos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimiento cobros"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "listado cobros"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimiento pagos"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listado pagos"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remesas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Transferencias"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listado tesoreria"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Integracion error"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar empresa"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar impresora seleccionada"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   6240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList11 
      Left            =   5520
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   8040
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpal.frx":34DE1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13494
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "12:38"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnDatos 
      Caption         =   "D&atos generales"
      Begin VB.Menu mnAgentes 
         Caption         =   "Agentes"
      End
      Begin VB.Menu mnDepartamentos 
         Caption         =   "Departamentos"
      End
      Begin VB.Menu mnFormasPago 
         Caption         =   "Formas de pago"
      End
      Begin VB.Menu mnTiposPago 
         Caption         =   "Tipos de pago"
      End
      Begin VB.Menu mnBancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu mnBIC 
         Caption         =   "BIC / SWIFT"
      End
      Begin VB.Menu mnbarra9 
         Caption         =   "-"
      End
      Begin VB.Menu mnConfiguracionAplicacion 
         Caption         =   "Confi&guracion"
         Begin VB.Menu mnParametros 
            Caption         =   "&Parametros"
         End
         Begin VB.Menu mnbarra15 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnMantenUsu 
            Caption         =   "Mantenimiento de usuarios"
         End
      End
      Begin VB.Menu mnCambioUsuario 
         Caption         =   "Cambiar  empresa"
      End
      Begin VB.Menu mnbarra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnSeleccionarImpresora 
         Caption         =   "Seleccionar impresora"
      End
      Begin VB.Menu mnbarra10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSal 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnCarteraCobros 
      Caption         =   "Cartera de cobros"
      Begin VB.Menu mnManteCobros 
         Caption         =   "Mantenimiento cobros"
      End
      Begin VB.Menu mnListadosCobros 
         Caption         =   "Cobros pendientes"
      End
      Begin VB.Menu mnImprimirCobros 
         Caption         =   "Informe cobros pendientes"
      End
      Begin VB.Menu mnImprimirRecibos 
         Caption         =   "Impresion recibos"
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnOrdenarCobros 
         Caption         =   "Realizar cobros por ..."
         Begin VB.Menu mnCobroConfirming 
            Caption         =   "&Confirming"
         End
         Begin VB.Menu mnCobroEfecti 
            Caption         =   "&Efectivo"
         End
         Begin VB.Menu mnCobroRecibo 
            Caption         =   "&Recibo bancario"
            Visible         =   0   'False
         End
         Begin VB.Menu mnCobroPagare 
            Caption         =   "&Pagaré"
         End
         Begin VB.Menu mnCobroTalon 
            Caption         =   "Ta&lón"
         End
         Begin VB.Menu mnCobroTransferencia 
            Caption         =   "&Transferencia"
         End
         Begin VB.Menu mnCobroTarjeta 
            Caption         =   "Ta&rjeta de crédito"
         End
      End
      Begin VB.Menu mnBarr9 
         Caption         =   "-"
      End
      Begin VB.Menu mnTransferenciasAbonos 
         Caption         =   "Transferencias abonos"
      End
      Begin VB.Menu mnBarra8 
         Caption         =   "-"
      End
      Begin VB.Menu mnCompensaciones 
         Caption         =   "Compensaciones"
      End
      Begin VB.Menu mnCompensaCliente 
         Caption         =   "Compensar cliente"
      End
      Begin VB.Menu mnbarra5 
         Caption         =   "-"
      End
      Begin VB.Menu mnMatenimientoCartas 
         Caption         =   "Carta reclamación"
      End
      Begin VB.Menu mnEfectuarReclama 
         Caption         =   "Efectuar reclamacion"
      End
      Begin VB.Menu mnManteniReclamas 
         Caption         =   "Historico reclamaciones"
      End
      Begin VB.Menu mnNorma57 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnNorma57 
         Caption         =   "Norma 57.   Importar fichero"
         Index           =   1
      End
      Begin VB.Menu mnbarraRecaudacEjec 
         Caption         =   "-"
      End
      Begin VB.Menu mnRecaudacionEjecutiva 
         Caption         =   "Recaudacion ejecutiva"
         Begin VB.Menu mnRecaudacioEjecutiva1 
            Caption         =   "Generar datos"
         End
      End
   End
   Begin VB.Menu mnMenuRemesas 
      Caption         =   "Remesas"
      Begin VB.Menu mnRemesas 
         Caption         =   "Remesas"
      End
      Begin VB.Menu mnBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnRemCancelaCliente 
         Caption         =   "Cancelación cliente"
      End
      Begin VB.Menu mnRemConfirmacion 
         Caption         =   "Confirmación remesa"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra13 
         Caption         =   "-"
      End
      Begin VB.Menu mnContabilizarRemesa 
         Caption         =   "Abono remesa"
      End
      Begin VB.Menu mnBarra12 
         Caption         =   "-"
      End
      Begin VB.Menu mnMenuDevolucionRem 
         Caption         =   "Devolucion"
         Begin VB.Menu mnDevolucionRemesa 
            Caption         =   "Manual"
         End
         Begin VB.Menu mnDevolRemFichBanc 
            Caption         =   "Fichero bancario"
         End
         Begin VB.Menu mnDevolRemDesdeVto 
            Caption         =   "Desde vencimiento"
         End
      End
      Begin VB.Menu mnEliminarEfectos 
         Caption         =   "Eliminar riesgo"
      End
      Begin VB.Menu mnListadoImpagados 
         Caption         =   "Listado impagados"
      End
   End
   Begin VB.Menu mnTalonesPagares 
      Caption         =   "Talones y pagarés"
      Begin VB.Menu mnRecepDoc 
         Caption         =   "Recepcion documentos"
         Index           =   0
      End
      Begin VB.Menu mnRecepDoc 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "Remesas talones y pagarés"
         Index           =   0
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "Cancelación cliente"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "Abono remesa"
         Index           =   3
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "Devolucion"
         Index           =   5
      End
      Begin VB.Menu mnTalonesPagares1 
         Caption         =   "Eliminar riesgo"
         Index           =   6
      End
   End
   Begin VB.Menu mnCarteraPagos 
      Caption         =   "Cartera de Pagos"
      Begin VB.Menu mnMantePagos 
         Caption         =   "Mantenimiento pagos"
      End
      Begin VB.Menu mnListadoPagos 
         Caption         =   "Pagos pendientes"
      End
      Begin VB.Menu mnImprimirPagos 
         Caption         =   "Informe pagos pendientes"
      End
      Begin VB.Menu mnListadoPagosBanco 
         Caption         =   "Listado pagos banco"
      End
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mngastosFijos 
         Caption         =   "Gastos fijos"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEfetuarPago 
         Caption         =   "Transferencias"
      End
      Begin VB.Menu mnPagosDom 
         Caption         =   "Pagos domiciliados"
      End
      Begin VB.Menu mnbarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnOrdenarPagos 
         Caption         =   "Realizar pagos por ..."
         Begin VB.Menu mnPagosConfirming 
            Caption         =   "&Confirming"
         End
         Begin VB.Menu mnPagosEfectivo 
            Caption         =   "&Efectivo"
         End
         Begin VB.Menu mnPagosRecibo 
            Caption         =   "&Recibo bancario"
         End
         Begin VB.Menu mnPagosTalon 
            Caption         =   "&Talón"
         End
         Begin VB.Menu mnPagosPagare 
            Caption         =   "&Pagaré"
         End
         Begin VB.Menu mnPagosTransferencia 
            Caption         =   "Trans&ferencia"
            Visible         =   0   'False
         End
         Begin VB.Menu mnPagosTarjeta 
            Caption         =   "Tar&jeta de crédito"
         End
      End
   End
   Begin VB.Menu mnOpAseguradas 
      Caption         =   "Operaciones aseguradas"
      Begin VB.Menu mnAseg_Basicos 
         Caption         =   "Clientes asegurados"
      End
      Begin VB.Menu mnAseg_LisFacturacion 
         Caption         =   "Listado vencimientos de asegurados"
      End
      Begin VB.Menu mnAseg_Impagos 
         Caption         =   "Impagados"
      End
      Begin VB.Menu mnAseg_Efectos 
         Caption         =   "Efectos por asegurados"
      End
      Begin VB.Menu mnAseg_barra 
         Caption         =   "-"
      End
      Begin VB.Menu mnAseg_AvisosAseguradora 
         Caption         =   "Listado avisos aseguradora"
      End
      Begin VB.Menu mnListAsegVarios 
         Caption         =   "Informe comunicación al seguro"
         Index           =   0
      End
      Begin VB.Menu mnListAsegVarios 
         Caption         =   "Facturas pendientes"
         Index           =   1
      End
      Begin VB.Menu mnAseg_Comprobar 
         Caption         =   "Comprobar vencimientos operaciones aseguradas"
      End
   End
   Begin VB.Menu mnCaja 
      Caption         =   "Caja"
      Begin VB.Menu mnCobrosPagosCaja 
         Caption         =   "Introducción caja"
      End
      Begin VB.Menu mnMantenimientoCaja 
         Caption         =   "Movimientos caja"
         Visible         =   0   'False
      End
      Begin VB.Menu mnListadoCobrosPagosCaja 
         Caption         =   "Listado caja"
      End
      Begin VB.Menu mnCierreCaja 
         Caption         =   "Cierre caja"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra20 
         Caption         =   "-"
      End
      Begin VB.Menu mnUsuariosCaja 
         Caption         =   "Usuarios caja"
      End
   End
   Begin VB.Menu mnInformes 
      Caption         =   "&Informes"
      Begin VB.Menu mnDeudaAgrupada 
         Caption         =   "Informe situación por NIF"
      End
      Begin VB.Menu mnDesdeHataDeudaNIF 
         Caption         =   "Informe situación por cuenta"
      End
      Begin VB.Menu mnbarra11 
         Caption         =   "-"
      End
      Begin VB.Menu mnListadoPrevisional 
         Caption         =   "Informe tesorería"
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnBackUp 
         Caption         =   "Copia seguridad local"
      End
      Begin VB.Menu mnUsuariosActivos 
         Caption         =   "Usuarios activos"
      End
      Begin VB.Menu mnListadoCobrosAgentesLin 
         Caption         =   "Listado cobros agentes"
      End
      Begin VB.Menu mnComprobarIntegraciones 
         Caption         =   "Comprobar integraciones"
      End
   End
   Begin VB.Menu mnPuntoFinal 
      Caption         =   "&Soporte"
      Begin VB.Menu mnAyuda 
         Caption         =   "Ayuda"
      End
      Begin VB.Menu mnbarra7_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEnviarMail 
         Caption         =   "Enviar Mail"
      End
      Begin VB.Menu mnWeb 
         Caption         =   "Web Ariadna Software"
      End
      Begin VB.Menu mnCheckVersion 
         Caption         =   "Comprobar version operativa"
      End
      Begin VB.Menu mnBarra7_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnAcercaDE 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Dim PrimeraVez As Boolean
Private TieneEditorDeMenus As Boolean

Private Sub MDIForm_Activate()
Dim C As String
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
        
        ValoresIntegraciones 0  'TODO
        
        If Not vParam Is Nothing Then
            If vParam.ComprobarAlInicio Then
                Screen.MousePointer = vbHourglass
                
                CadenaDesdeOtroForm = "cuentas,tiposituacionrem,ctabancaria,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo "
                C = "remesas.codmacta=cuentas.codmacta and situacio=situacion and ctabancaria.codmacta=remesas.codmacta"
                C = C & " AND (situacion ='Q' or situacion ='Y') AND 1"
                C = DevuelveDesdeBD("count(*)", CadenaDesdeOtroForm, C, "1")
                If C = "" Then C = "0"
                If Val(C) > 0 Then
                    frmVarios.SubTipo = 3 'Las dos
                    frmVarios.Opcion = 12
                    frmVarios.Show vbModal
                End If
                
                
                'Si tienen operaciones aseguradas
                ComprobarOperacionesAseguradas True
                
            End If
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub ValoresIntegraciones(Opcion As Byte)
Dim Bol As Boolean
Dim HayErrores As Boolean
Dim SQL As String


    If (vParam Is Nothing) Or (vEmpresa Is Nothing) Then Exit Sub

    HayErrores = False
    If Opcion = 0 Or Opcion = 1 Then

    End If
    HayErrores = HayErrores Or Bol
    
    If Opcion = 0 Or Opcion = 2 Then
        'Buscamos las integraciones que psaron
        Bol = BuscarIntegraciones(True, Format(vEmpresa.codempre, "00"))
        Me.Toolbar11.Buttons(17).Visible = Bol
    End If
    HayErrores = HayErrores Or Bol
    
    SQL = ""
    
    
    If HayErrores And Opcion = 0 Then
        'Mostraremos el formulario
        frmintegraciones.TablasDeErrores = SQL
        frmintegraciones.Show vbModal
    End If
    
    ComprobarFuncionamientoEspia
    
End Sub


Private Sub MDIForm_Load()
    PrimeraVez = True
    
    ImageList11.ImageHeight = 24
    ImageList11.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\imgppalmon.dll", 1, 24
    imgListComun.ImageHeight = 24
    imgListComun.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\icocommon.dll", 2, 24
    
    
    'Botones
    With Me.Toolbar11
        
        .ImageList = Me.ImageList11
        .Buttons(1).Image = 13
        '---
        .Buttons(2).Image = 19
        .Buttons(3).Image = 18
        
        .Buttons(6).Image = 14
        .Buttons(7).Image = 16
        '----
        .Buttons(9).Image = 15
        .Buttons(10).Image = 17
    
        .Buttons(12).Image = 11

        .Buttons(13).Image = 12
        
        .Buttons(15).Image = 10

        

'        .Buttons(14).Image = 17   'Cuenta P y G
        .Buttons(17).Image = 9
     '   .Buttons(17).Visible = TieneIntegracionesPendientes
        
        
        '----
        .Buttons(19).Image = 2  'Usuarios
        .Buttons(20).Image = 1  'Impresora
        .Buttons(21).Image = 3  'Salir

    End With
    LeerEditorMenus
    PonerDatosFormulario
       
    mnComprobarIntegraciones = vConfig.Integraciones <> ""
    
End Sub


Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vParam Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True
    
    
    If vEmpresa Is Nothing Then
        CadenaDesdeOtroForm = "Configurando"
    Else
        CadenaDesdeOtroForm = vEmpresa.nomempre & " (" & vEmpresa.codempre & ")"
    End If
    Me.StatusBar1.Panels(2).Text = CadenaDesdeOtroForm
    
    'FijarConerrores
    CadenaDesdeOtroForm = ""
    
    'Poner datos visible del form
    PonerDatosVisiblesForm
    'Poner opciones de nivel de usuario
    PonerOpcionesUsuario
    
    'Si tiene editor de menus
    mnOpAseguradas.Visible = True
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
    If mnOpAseguradas.Visible Then
        mnOpAseguradas.Visible = vParam.TieneOperacionesAseguradas
        mnAseg_Efectos.Visible = False 'De momento siempre oculto
    End If
    
    If mnbarraRecaudacEjec.Visible Then
        mnbarraRecaudacEjec.Visible = vParam.RecaudacionEjecutiva
        mnbarraRecaudacEjec.Visible = vParam.RecaudacionEjecutiva
    End If
    If mnRecaudacionEjecutiva.Visible Then
        mnRecaudacionEjecutiva.Visible = vParam.RecaudacionEjecutiva
        mnRecaudacionEjecutiva.Visible = vParam.RecaudacionEjecutiva
    End If
    
    If vParam.Norma57 = 0 Then
        Me.mnNorma57(0).Visible = False
        Me.mnNorma57(1).Visible = False
    End If
    
    If vParam.PagosConfirmingCaixa Then
        'Los confirming los pagara por NORMA banacaria
        mnPagosDom.Caption = "Caixa confirming"
    Else
        mnPagosDom.Caption = "Pagos domiciliados"
    End If
    
    
    'ASociacion entre botones y menus
      With Me.Toolbar11
        
        .Buttons(1).Visible = mnDatos.Visible And Me.mnFormasPago.Visible
        .Buttons(2).Visible = mnDatos.Visible And Me.mnTiposPago.Visible
        .Buttons(3).Visible = mnDatos.Visible And Me.mnBancos.Visible
        '---
        .Buttons(6).Visible = mnCarteraCobros.Visible And Me.mnManteCobros.Visible    '
        .Buttons(7).Visible = mnCarteraCobros.Visible And mnImprimirCobros.Visible    '
     
        '----
        .Buttons(9).Visible = mnCarteraPagos.Visible And Me.mnMantePagos.Visible
        .Buttons(10).Visible = mnCarteraPagos.Visible And Me.mnImprimirPagos.Visible
        
        '----
        .Buttons(12).Visible = mnMenuRemesas.Visible And mnRemesas.Visible  'Balance
        .Buttons(13).Visible = mnCarteraPagos.Visible And mnEfetuarPago.Visible  'Cuenta P y G
        .Buttons(15).Visible = mnInformes.Visible And mnListadoPrevisional.Visible
        
        
    End With
    
    
    'Me.mnListadoPagosBanco.Visible = True  'deberia hacerlo por esta en el preoceso pagos-ordenar pagos prov-> efectos  LISTADO BANCO
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim Cad As String


    'Alguna cosilla antes de cerrar. Eliminar bloqueos
    Cad = "Delete from zBloqueos where codusu = " & vUsu.Codigo
    Conn.Execute Cad

    'Elimnar bloquo BD
    Cad = DevuelveDesdeBD("codusu", "Usuarios.vBloqBD", "codusu", vUsu.Codigo)
    If Cad <> "" Then Conn.Execute "Delete from Usuarios.vBloqBD where codusu=" & vUsu.Codigo
        
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub mnAcercaDE_Click()
    frmVarios.Opcion = 11
    frmVarios.Show vbModal
End Sub

Private Sub mnAgentes_Click()

    frmAgentes.Show vbModal
End Sub

Private Sub mnAseg_AvisosAseguradora_Click()
    frmListado.Opcion = 33
    frmListado.Show vbModal
End Sub

Private Sub mnAseg_Basicos_Click()
    frmListado.Opcion = 15
    frmListado.Show vbModal
End Sub

Private Sub mnAseg_Comprobar_Click()
    ComprobarOperacionesAseguradas False
End Sub

Private Sub mnAseg_Efectos_Click()
    frmListado.Opcion = 18
    frmListado.Show vbModal
End Sub

Private Sub mnAseg_Impagos_Click()
    frmListado.Opcion = 17
    frmListado.Show vbModal
End Sub

Private Sub mnAseg_LisFacturacion_Click()
    frmListado.Opcion = 16
    frmListado.Show vbModal
End Sub

Private Sub mnBackUp_Click()
    frmBackUP.Show vbModal
End Sub

Private Sub mnBancos_Click()
    frmBanco.Show vbModal
End Sub

Private Sub mnBIC_Click()
    frmbic.Show vbModal
End Sub

Private Sub mnCambioUsuario_Click()

    
    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If
    
    'Borramos temporal
    Conn.Execute "Delete from zBloqueos where codusu = " & vUsu.Codigo

    
    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"

    frmLogin.Show vbModal

    
    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    Conn.Close

    
    If AbrirConexion() = False Then
        MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If
    
    
    Set vParam = Nothing
    Set vEmpresa = Nothing
    
    
    'Los ponemos a true, y despues se ajustaran
    mnbarraRecaudacEjec.Visible = True
    mnOpAseguradas.Visible = True
    mnRecaudacionEjecutiva.Visible = True
    Me.mnNorma57(0).Visible = True
    Me.mnNorma57(1).Visible = True
    
    LeerEmpresaParametros
    PonerDatosFormulario
    
    'Ponemos primera vez a false
    PrimeraVez = True
    Me.SetFocus
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnCheckVersion_Click()
   Screen.MousePointer = vbHourglass
    LanzaHome "webversion"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

'Private Sub mnCierreCaja_Click()
'
'    MsgBox "No disponible"
'
''    frmCaja.Opcion = 2
''    'frmCaja.LaCuentaDeCaja = ObtenerCuentaCaja(False)
''    'frmCaja.LaCuentaDeCaja = ObtenerCuentaCaja(True)
''    If frmCaja.LaCuentaDeCaja = "" Then
''        MsgBox "El usuario: " & vUsu.Nombre & " NO tiene cuenta de caja asignada", vbExclamation
''    Else
''        frmCaja.Show vbModal
''    End If
'End Sub

Private Sub mnCobroConfirming_Click()
    Cobros 5
End Sub


Private Sub Cobros(Tipo As Byte)
    frmVarios.Opcion = 0
    frmVarios.SubTipo = Tipo
    frmVarios.Show vbModal
End Sub

Private Sub Pagos(Tipo As Byte)
    frmVarios.Opcion = 1
    frmVarios.SubTipo = Tipo
    frmVarios.Show vbModal
End Sub

Private Sub mnCobroEfecti_Click()
    Cobros 0
End Sub

Private Sub mnCobroPagare_Click()
    Cobros 3
End Sub


Private Sub mnCobrosPagosCaja_Click()
Dim CajaPpal As Boolean
    If ObtenerCuentaCaja2(True, CajaPpal) = "" Then
        MsgBox "El usuario: " & vUsu.Nombre & " NO tiene cuenta de caja asignada", vbExclamation
    Else
        frmCajaN.UsuarioCajaPredeterminada = CajaPpal
        frmCajaN.Show vbModal
    End If
    
    
End Sub

Private Sub mnCobroTalon_Click()
    Cobros 2
End Sub

Private Sub mnCobroTarjeta_Click()
    Cobros 6
End Sub

Private Sub mnCobroTransferencia_Click()
    Cobros 1
End Sub

Private Sub mnCompensaciones_Click()
    frmCompensaciones.Show vbModal
End Sub

Private Sub mnCompensaCliente_Click()
    CadenaDesdeOtroForm = ""
    frmListado.Opcion = 36
    frmListado.Show vbModal
End Sub

Private Sub mnComprobarIntegraciones_Click()
    Me.Toolbar11.Buttons(17).Visible = False
    ValoresIntegraciones 0
End Sub

Private Sub mnContabilizarRemesa_Click()
    'Antiguamente llamado CONTABILIZACION REMESA
    frmVarios.SubTipo = 1
    frmVarios.Opcion = 8
    frmVarios.Show vbModal
End Sub

Private Sub mnDepartamentos_Click()
    frmDepartamentos.Show vbModal
End Sub

Private Sub mnDesdeHataDeudaNIF_Click()
    frmVarios.Opcion = 14
    frmVarios.Show vbModal
End Sub

Private Sub mnDeudaAgrupada_Click()
    frmVarios.Opcion = 13
    frmVarios.Show vbModal
End Sub

Private Sub mnDevolRemDesdeVto_Click()
    frmVarios.SubTipo = 1
    frmVarios.Opcion = 28
    frmVarios.Show vbModal
End Sub

Private Sub mnDevolRemFichBanc_Click()
    frmVarios.SubTipo = 1
    frmVarios.Opcion = 16
    frmVarios.Show vbModal
End Sub

Private Sub mnDevolucionRemesa_Click()
    frmVarios.SubTipo = 1
    frmVarios.Opcion = 9
    frmVarios.Show vbModal
End Sub

Private Sub mnEfectuarReclama_Click()
    frmListado.Opcion = 3
    frmListado.Show vbModal
End Sub

Private Sub mnEfetuarPago_Click()
    'Pagos 1
    frmTransferencias2.TipoDeFrm = 1  'TRANSFERENCIA
    frmTransferencias2.Show vbModal
End Sub

Private Sub mnEliminarEfectos_Click()
    frmVarios.SubTipo = 1 'EFECTOS
    frmVarios.Opcion = 12
    frmVarios.Show vbModal
End Sub


Private Sub mnEnviarMail_Click()
    frmEMail.Opcion = 1
    frmEMail.Show vbModal
End Sub


Private Sub mnFormasPago_Click()
    frmFormaPago.Show vbModal
End Sub


Private Sub mngastosFijos_Click()
    frmGastosFijos.Show vbModal
End Sub

Private Sub mnImprimirCobros_Click()
    frmListado.Opcion = 1
    frmListado.Show vbModal
End Sub

Private Sub mnImprimirPagos_Click()
    frmListado.Opcion = 2
    frmListado.Show vbModal
End Sub

Private Sub mnImprimirRecibos_Click()
    frmVarios.Opcion = 24
    frmVarios.Show vbModal
    
End Sub

Private Sub mnListadoCobrosAgentesLin_Click()
    frmListado.Opcion = 45
    frmListado.Show vbModal
End Sub

Private Sub mnListadoCobrosPagosCaja_Click()
    frmListado.Opcion = 8
    frmListado.Show vbModal
End Sub

Private Sub mnListadoImpagados_Click()
    frmListado.Opcion = 9
    frmListado.Show vbModal
End Sub

Private Sub mnListadoPagos_Click()
    frmVerCobrosPagos.vSQL = ""
    frmVerCobrosPagos.OrdenarEfecto = False
    frmVerCobrosPagos.Regresar = False
    frmVerCobrosPagos.Cobros = False
    frmVerCobrosPagos.Show vbModal
End Sub

Private Sub mnListadoPagosBanco_Click()
    frmListado.Opcion = 25
    frmListado.Show vbModal
End Sub

Private Sub mnListadoPrevisional_Click()
    frmListado.Opcion = 12
    frmListado.Show vbModal
End Sub

Private Sub mnListadosCobros_Click()
    frmVerCobrosPagos.vSQL = ""
    frmVerCobrosPagos.OrdenarEfecto = False
    frmVerCobrosPagos.Regresar = False
    frmVerCobrosPagos.Cobros = True
    frmVerCobrosPagos.Show vbModal
End Sub

Private Sub mnListAsegVarios_Click(Index As Integer)
    If Index = 0 Then
        frmListado.Opcion = 39
    Else
        frmListado.Opcion = 40
    End If
    
    frmListado.Show vbModal
    
End Sub

Private Sub mnManteCobros_Click()
    frmCobros.Show vbModal
End Sub

'Private Sub mnMantenimientoCaja_Click()
'    frmCaja.Opcion = 1
'    'frmCaja.LaCuentaDeCaja = ObtenerCuentaCaja(False)
'    If frmCaja.LaCuentaDeCaja = "" Then
'        MsgBox "El usuario: " & vUsu.Nombre & " NO tiene cuenta de caja asignada", vbExclamation
'    Else
'        frmCaja.Show vbModal
'    End If
'End Sub

Private Sub mnManteniReclamas_Click()
    frmReclama.Show vbModal
    Exit Sub
    frmColReclamas.Show vbModal
End Sub

Private Sub mnMantenUsu_Click()
    frmMantenusu.Show vbModal
End Sub

Private Sub mnMantePagos_Click()
    frmPagoPro.Show vbModal
End Sub

Private Sub mnMatenimientoCartas_Click()
    frmFacCartasOferta.Show vbModal
End Sub


Private Sub mnNorma57_Click(Index As Integer)
    If Index = 1 Then
        frmListado.Opcion = 42
        frmListado.Show vbModal
    End If
End Sub



Private Sub mnPagosConfirming_Click()
    Pagos 5
End Sub

Private Sub mnPagosDom_Click()
    frmTransferencias2.TipoDeFrm = 2  'TRANSFERENCIA
    frmTransferencias2.Show vbModal
End Sub

Private Sub mnPagosEfectivo_Click()
    Pagos 0
End Sub

Private Sub mnPagosPagare_Click()
    Pagos 3
End Sub

Private Sub mnPagosRecibo_Click()
    Pagos 4
End Sub
    '
Private Sub mnPagosTalon_Click()
    Pagos 2
End Sub

Private Sub mnPagosTarjeta_Click()
    Pagos 6
End Sub

Private Sub mnPagosTransferencia_Click()
    Pagos 1
End Sub

Private Sub mnParametros_Click()
    frmparametros.Show
End Sub

Private Sub mnRecaudacioEjecutiva1_Click()
    frmListado.Opcion = 38
    frmListado.Show vbModal
End Sub

Private Sub mnRecepDoc_Click(Index As Integer)
    frmRecpcionDoc.Show vbModal
End Sub

Private Sub mnRemCancelaCliente_Click()
    If Not vParam.EfectosCtaPuente Then
        MsgBox "Falta configurar en parametros", vbExclamation
        Exit Sub
    End If

    frmVarios.SubTipo = 1
    frmVarios.Opcion = 22
    frmVarios.Show vbModal
End Sub

Private Sub mnRemConfirmacion_Click()
    frmVarios.SubTipo = 1
    frmVarios.Opcion = 23
    frmVarios.Show vbModal
End Sub

Private Sub mnRemesas_Click()
    frmColRemesas2.Tipo = 1 'EFECTOS
    frmColRemesas2.Show vbModal
End Sub

Private Sub mnSeleccionarImpresora_Click()
    Me.CommonDialog1.Flags = cdlPDPrintSetup
    Me.CommonDialog1.ShowPrinter
End Sub


Private Sub mnTalonesPagares1_Click(Index As Integer)
    Select Case Index
    Case 0
            'Mantenimiento remesas
            frmColRemesas2.Tipo = 2
            frmColRemesas2.Show vbModal
    Case 2
            'Cancelacion
            frmVarios.SubTipo = 2
            frmVarios.Opcion = 22
            frmVarios.Show vbModal
            
    Case 3
            'Abono
            frmVarios.SubTipo = 2
            frmVarios.Opcion = 8
            frmVarios.Show vbModal
            
    Case 5
            'Devolucion
            frmVarios.SubTipo = 2
            frmVarios.Opcion = 9
            frmVarios.Show vbModal
            
            
    Case 6
            frmVarios.SubTipo = 2
            frmVarios.Opcion = 12
            frmVarios.Show vbModal
    End Select
End Sub



Private Sub mnTiposPago_Click()
    frmTipoPago.Show vbModal
End Sub

Private Sub mnTransferenciasAbonos_Click()
    frmTransferencias2.TipoDeFrm = 0   'ABONOS
    frmTransferencias2.Show vbModal
End Sub

Private Sub mnuSal_Click()
    Unload Me
End Sub

Private Sub mnUsuariosActivos_Click()
Dim SQL As String
Dim I As Integer
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        I = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, I)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            I = I + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningún usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub
'

Private Sub mnUsuariosCaja_Click()
    frmUsuariosCaja.Show vbModal
End Sub

Private Sub mnWeb_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerDatosVisiblesForm()
Dim Cad As String
    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    Cad = Cad & ", " & Format(Now, "d")
    Cad = Cad & " de " & Format(Now, "mmmm")
    Cad = Cad & " de " & Format(Now, "yyyy")
    Cad = "    " & Cad & "    "
    Me.StatusBar1.Panels(5).Text = Cad
    'Caption = "ARIMONEY" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  "
    Caption = "ARIMONEY  "
    If vEmpresa Is Nothing Then
        Caption = Caption & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
    Else
        Caption = Caption & "Ver:  " & App.Major & "." & App.Minor & "." & App.Revision & "   " & vEmpresa.nomresum & "    Usuario: " & vUsu.Nombre
    End If
End Sub




Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Name, 1, 6)) <> "mnbarr" Then _
                T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar11.Enabled = Habilitar
    Me.Toolbar11.Visible = Habilitar
    mnParametros.Enabled = True
    'mnEmpresa.Enabled = True
    Me.mnParametros.Enabled = True
    Me.mnConfiguracionAplicacion.Enabled = True
    mnDatos.Enabled = True
    Me.mnuSal.Enabled = True
    Me.mnCambioUsuario.Enabled = True
End Sub



Private Sub PonerOpcionesUsuario()
    Dim B As Boolean

    B = vUsu.Nivel < 2  'Administradores y root
    mnMantenUsu.Enabled = B
        
    B = vUsu.Nivel = 3  'Es usuario de consultas
    If B Then
    
    End If
End Sub



Private Sub LanzaHome(Opcion As String)
    Dim I As Integer
    Dim Cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD(Opcion, "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    If Opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "/tesor.cfm?version=" & App.Major & "." & App.Minor & "." & App.Revision
        
    I = FreeFile
    Cad = ""
    If Dir(App.Path & "\lanzaexp.dat", vbArchive) = "" Then
        MsgBox "Faltan archivo de configuracion: lanzaexp.dat", vbExclamation
    Else
        
        Open App.Path & "\lanzaexp.dat" For Input As #I
        Line Input #I, Cad
        Close #I
    End If
    'Lanzamos
    If Cad <> "" Then Shell Cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, Cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub

Private Sub HacerToolBar(ToolBarOpcion As Integer)
    Select Case ToolBarOpcion
    Case 1
        mnTiposPago_Click
            
    Case 2
        mnFormasPago_Click
    
    Case 3
         mnBancos_Click
        
        
    Case 6
        mnManteCobros_Click
        
    Case 7
        mnImprimirCobros_Click
        
    Case 9
        mnMantePagos_Click
        
    Case 10
        mnImprimirPagos_Click
        
    Case 12
        mnRemesas_Click
        
    Case 13
        'transferencias
         mnEfetuarPago_Click
         
    Case 15
        mnListadoPrevisional_Click
        
    Case 17
        frmintegraciones.TablasDeErrores = ""
        frmintegraciones.Show vbModal
        
    Case 19
        mnCambioUsuario_Click
        
    Case 20
        Screen.MousePointer = vbHourglass
        Me.CommonDialog1.ShowPrinter
        Screen.MousePointer = vbDefault
        
    Case 21
        Unload Me
    End Select
End Sub


 

Private Sub Toolbar11_ButtonClick(ByVal Button As MSComctlLib.Button)


    
    HacerToolBar Button.Index
End Sub


Private Function ObtenerCuentaCaja2(HacerPagosCobros As Boolean, ByRef Principal As Boolean) As String
Dim SQL As String

    ObtenerCuentaCaja2 = ""
    SQL = "Select  susucaja.*, nommacta "
    SQL = SQL & " from susucaja,cuentas "
    SQL = SQL & " WHERE ctacaja = cuentas.codmacta"
    SQL = SQL & " AND codusu =" & Val(Right(CStr(vUsu.Codigo), 2))
    Principal = False
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Principal = (miRsAux!predeterminado = 1)
        If HacerPagosCobros Then
            'Es para efectura los pagos.
            ObtenerCuentaCaja2 = miRsAux!CtaCaja
        Else
            'ES para ver el mantenimiento de cobros / pagos
            'Indicaremos si el usuario tiene cta o si es super usuario caja
            If miRsAux!predeterminado = 1 Then
                ObtenerCuentaCaja2 = "S"
            Else
                ObtenerCuentaCaja2 = miRsAux!CtaCaja
            End If
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function


'''ICONOS
Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim I As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub



Private Sub LeerEditorMenus()
Dim SQL As String
    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='Tesor'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='Tesor' and codusu = " & Val(Right(CStr(vUsu.Codigo), 2))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is Menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, SQL, C) > 0 Then T.Visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index & "|"
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function


Private Sub ComprobarOperacionesAseguradas(AlInicio As Boolean)

        If vParam.TieneOperacionesAseguradas Then
            If vUsu.Nivel = 0 Then
                NumRegElim = 0
                'Avisos falta
                If Asegurados_HayAvisos(True) Then NumRegElim = 1
                'Siniestros
                If Asegurados_HayAvisos(False) Then NumRegElim = NumRegElim + 2
                    
                If NumRegElim > 0 Then
                    If NumRegElim > 2 Then
                        frmListadoASegurado.Opcion = 0
                    Else
                        'Solo uno de los dos
                        frmListadoASegurado.Opcion = CByte(NumRegElim)
                    End If
                    frmListadoASegurado.Show vbModal
                
                Else
                    If Not AlInicio Then MsgBox "Ningun valor devuelto", vbInformation
                End If
            End If
        End If
End Sub

