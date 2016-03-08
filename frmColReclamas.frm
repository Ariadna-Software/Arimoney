VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmColReclamas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reclamaciones efectuadas"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "frmColReclamas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColReclamas.frx":000C
      Height          =   5925
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   10451
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
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6000
      Top             =   5640
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
      TabIndex        =   1
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmColReclamas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SQL As String

Private Sub Form_Load()


      With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 2
        .Buttons(4).Image = 4
        .Buttons(5).Image = 5
        .Buttons(7).Image = 16
        .Buttons(9).Image = 15

    End With
    Me.Icon = frmPpal.Icon
    HacerToolBar 1  'Ver todos

End Sub



Private Sub CargaGrid()
Adodc1.ConnectionString = Conn

    SQL = "SELECT fecreclama,codmacta,nommacta,if(carta=0,'*',' '),impvenci"
    SQL = SQL & ",fecfaccl,numserie,codfaccl,numorden,codigo"
    SQL = SQL & " FROM shcocob"
    SQL = SQL & " ORDER BY fecreclama,fecfaccl,codmacta"
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    
   
        DataGrid1.Columns(0).Caption = "Reclama"
        DataGrid1.Columns(0).Width = 1200
    
   
        DataGrid1.Columns(1).Caption = "Cuenta"
        DataGrid1.Columns(1).Width = 1000

    
   
        DataGrid1.Columns(2).Caption = "Denominación"
        DataGrid1.Columns(2).Width = 2500

        DataGrid1.Columns(3).Caption = "@"
        DataGrid1.Columns(3).Width = 500


'SQL = "SELECT fecreclama,codmacta,nommacta,carta,impvenci"
'SQL = SQL & ",fecfaccl,numserie,codfaccl,numorden,codigo"
        
        DataGrid1.Columns(4).Caption = "Importe"
        DataGrid1.Columns(4).Width = 1000
        DataGrid1.Columns(4).NumberFormat = FormatoImporte
        DataGrid1.Columns(4).Alignment = dbgRight

        DataGrid1.Columns(5).Caption = "F. Factura"
        DataGrid1.Columns(5).Width = 1200
        DataGrid1.Columns(5).Alignment = dbgCenter

        DataGrid1.Columns(6).Caption = "serie"
        DataGrid1.Columns(6).Width = 650

        DataGrid1.Columns(7).Caption = "Codigo"
        DataGrid1.Columns(7).Width = 1100
        DataGrid1.Columns(7).Alignment = dbgRight

        DataGrid1.Columns(8).Caption = "Vto."
        DataGrid1.Columns(8).Width = 450

        'OCulto el codigo
        DataGrid1.Columns(9).Visible = False
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    HacerToolBar Button.Index
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerToolBar(button_index As Integer)
    Select Case button_index
    Case 1
        CargaGrid
        
    'Case 4,5 modificar,eliminar
    
    
    Case 7
        'IMPRIMIR
        
        MsgBox "Opcion NO disponible.", vbExclamation
    Case 9
        Unload Me
    End Select
End Sub
