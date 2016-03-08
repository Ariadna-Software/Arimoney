VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmBackUP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmBackUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl2.Animation Animation1 
      Height          =   915
      Left            =   300
      TabIndex        =   5
      Top             =   1500
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1614
      _Version        =   327681
      FullWidth       =   301
      FullHeight      =   61
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   1740
      TabIndex        =   3
      Top             =   3120
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmBackUP.frx":000C
      Left            =   120
      List            =   "frmBackUP.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "sobre ficheros locales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   4860
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
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Copia de seguridad :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   3480
   End
End
Attribute VB_Name = "frmBackUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    
'CONSTANTES
Private TesoreriaTablas As String
Private TesoreriaNumeroTablas As Byte


Private Tablas() As String
Private NumTablas As Integer

Dim RS As Recordset
Dim NF As Integer
Dim Archivo As String
Dim Izquierda As String
Dim Derecha As String


'En todas las futuros backups, se trata de cargar el array tablas con las tablas(-1) a copiar


Private Sub cmdAceptar_Click()
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de copia", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Select Case Combo1.ListIndex
    Case 0
        CopiaTodo
'    Case 1
'        RenumeracionAsientos
'    Case 2
'        Cierre
'    Case 3
'        Amorizacion
    End Select
    If NumTablas = 0 Then
        MsgBox "Error verificando numero de tablas", vbExclamation
        Exit Sub
    Else
        'Ahora hacemos las copias
        HacerBackUp
        MsgBox "Copia finalizada en: " & Archivo, vbInformation
        cmdAceptar.Enabled = False
        Label1.Caption = ""
        PonerVideo False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = ""
    Label2.Caption = "Empresa: " & vEmpresa.nomempre
    Caption = "Backup para " & UCase(vEmpresa.nomresum)
    Me.Icon = frmPpal.Icon
    
  
    
    'Tablas de TESORERIA
    '-------------------------------------------------
'    '                                       8 tablas
'    TesoreriaTablas = "scaja|shcaja|usucaja|sefecdev|ctabancaria|Departamentos|Agentes|scobro|"
'    '                                       6 tablas
'    TesoreriaTablas = TesoreriaTablas & "spagop|stipoformapago|scartas|Remesas|sforpa|"
'
    
    
    TesoreriaTablas = "agentes" & "|" & "paramtesor" & "|"
    TesoreriaTablas = TesoreriaTablas & "remesas" & "|"
    TesoreriaTablas = TesoreriaTablas & "scartas" & "|" & "Departamentos" & "|"
    TesoreriaTablas = TesoreriaTablas & "sefecdev" & "|" & "sforpa" & "|"
    TesoreriaTablas = TesoreriaTablas & "sgastfij" & "|" & "sgastfijd" & "|"
    TesoreriaTablas = TesoreriaTablas & "agentes" & "|" & "spagop" & "|"
    TesoreriaTablas = TesoreriaTablas & "stipoformapago" & "|" & "stransfer" & "|"
    TesoreriaTablas = TesoreriaTablas & "tiporemesa" & "|" & "tiposituacionrem" & "|"
    TesoreriaTablas = TesoreriaTablas & "susucaja" & "|" & "stransfercob" & "|"
    TesoreriaTablas = TesoreriaTablas & "shcocob" & "|" & "scobro" & "|"
    TesoreriaTablas = TesoreriaTablas & "slicaja" & "|" & "scacaja" & "|"
    
    TesoreriaNumeroTablas = 20
    CargaCombo
End Sub


Private Sub CargaCombo()
Combo1.Clear
Combo1.AddItem "Copia todo "
'Combo1.AddItem "Renumeración asientos"
'Combo1.AddItem "Cierre ejercicio"
'Combo1.AddItem "Amortización"
End Sub

Private Sub CopiaTodo()

    NumTablas = TesoreriaNumeroTablas
    ReDim Tablas(NumTablas - 1)
    For NF = 1 To NumTablas
        Izquierda = RecuperaValor(TesoreriaTablas, NF)
        If Izquierda = "" Then
            NumTablas = 0
            Exit For
        Else
            Tablas(NF - 1) = Izquierda
        End If
    Next NF
    
End Sub


Private Sub PonerVideo(Encender As Boolean)
    If Encender Then
        Me.Animation1.Open App.Path & "\actua.avi"
        Me.Animation1.Play
        Me.Animation1.Visible = True
    Else
        Me.Animation1.Stop
        Me.Animation1.Visible = False
    End If
End Sub



Private Sub HacerBackUp()
Dim I As Integer

    If NumTablas > 3 Then PonerVideo True


    Archivo = FijarCarpeta
    If Archivo = "" Then
        MsgBox "no se ha creado correctamente la carpeta de copia.", vbExclamation
        Exit Sub
    End If
        
    
    For I = 0 To NumTablas - 1
        Label1.Caption = Tablas(I) & "     (" & I + 1 & " de " & NumTablas & ")"
        Label1.Refresh
        BKTablas (Tablas(I))
    Next I
    
    
    
End Sub

'

Private Function FijarCarpeta() As String
Dim FE As String
Dim I As Integer

On Error GoTo EFijarCarpeta
    FijarCarpeta = ""
    If Dir(App.Path & "\BACKUP", vbDirectory) = "" Then MkDir App.Path & "\BACKUP"
    
    Derecha = App.Path & "\BACKUP\T"
    Izquierda = Format(Now, "yymmdd")
    I = -1
    Do
        I = I + 1
        FE = Format(I, "00")
        FE = Derecha & Izquierda & FE
        If Dir(FE, vbDirectory) = "" Then
            'OK
            MkDir FE
            FijarCarpeta = FE
            I = 100
        End If
    Loop Until I > 99
    Exit Function
EFijarCarpeta:
    MuestraError Err.Number
End Function



Private Sub BKTablas(Tabla As String)
Dim Cad As String
    Set RS = New ADODB.Recordset
    RS.Open Tabla, Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RS.EOF Then
        'No hace falta hacer back up
    
    Else
        NF = FreeFile
        Open Archivo & "\" & Tabla & ".sql" For Output As #NF
        BACKUP_TablaIzquierda RS, Izquierda
        While Not RS.EOF
            BACKUP_Tabla RS, Derecha
            Cad = "INSERT INTO " & Tabla & " " & Izquierda & " VALUES " & Derecha & ";"
            Print #NF, Cad
            RS.MoveNext
        Wend
        Close #NF
    End If
    RS.Close
    Set RS = Nothing
End Sub


'Private Sub RenumeracionAsientos()
'    NumTablas = 6
'    ReDim Tablas(NumTablas - 1)
'    Tablas(0) = "hcabapu"
'    Tablas(1) = "hlinapu"
'    Tablas(2) = "cabfact"
'    Tablas(3) = "cabfactprov"
'    Tablas(4) = "linfact"
'    Tablas(5) = "linfactprov"
'End Sub
'
'
'
'Private Sub Cierre()
'    NumTablas = 5
'    ReDim Tablas(NumTablas - 1)
'    Tablas(0) = "hcabapu"
'    Tablas(1) = "hlinapu"
'    Tablas(2) = "contadores"
'    Tablas(3) = "hsaldos"
'    Tablas(4) = "hsaldosanal"
'
'End Sub
'
'Private Sub Amorizacion()
'    NumTablas = 3
'    ReDim Tablas(NumTablas - 1)
'    Tablas(0) = "samort"
'    Tablas(1) = "sinmov"
'    Tablas(2) = "shisin"
'End Sub
'
'
