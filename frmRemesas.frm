VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRemesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesas"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12525
   Icon            =   "frmRemesas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   12525
   Begin VB.Frame FrameRemesas 
      Height          =   2910
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   6360
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1665
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4725
         TabIndex        =   4
         Top             =   1170
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   3
         Top             =   2205
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   2
         Top             =   2205
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Abono"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   90
         TabIndex        =   8
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Left            =   1575
         Picture         =   "frmRemesas.frx":030A
         Top             =   1665
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Remesa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   7
         Top             =   1215
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Diskette Remesa Norma 19"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   855
         TabIndex        =   6
         Top             =   405
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   810
         TabIndex        =   5
         Top             =   1215
         Width           =   690
      End
      Begin VB.Image ImgRem 
         Height          =   240
         Index           =   0
         Left            =   1575
         Picture         =   "frmRemesas.frx":040C
         Top             =   1192
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6900
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0-Diskette Norma 19
    '1-Eliminar remesa
    '2-Abono de remesa
    
Public Event DatoSeleccionado(CadenaSeleccion As String)

Dim Tablas As String

Private WithEvents frmM As frmMensajes
Attribute frmM.VB_VarHelpID = -1

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim Cad1 As String
Dim Cont As Long
Dim i As Integer

Dim Importe As Currency

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
    
    If Text1(0).Text = "" Then
        MsgBox " Debe de introducir los datos de la Remesa previamente.", vbExclamation
        Exit Sub
    End If
    
    'Hacemos el select y si tiene resultados mostramos los valores
    Cad = " SELECT sremes.* from sremes WHERE numremes =  " & Text1(0).Text

    If Opcion = 1 Or Opcion = 2 Then Cad = Cad & " and situacio = 0"

    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RS.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
    Else
        'Mostramos el frame de resultados
        'Text3(7).Text = RS!fecremes
        
        Select Case Opcion
        Case 0  ' generacion de diskette norma 19
            Dim nomfich As String
            ' aqui tenemos que grabar el diskette
            nomfich = InputBox("Nombre del fichero a generar:", "Generación de remesas", "A:\Remesas.txt")
            If nomfich = "" Then Exit Sub ' Le han dado a cancelar
            If GrabarDisketteNorma19(nomfich, CLng(Text1(0).Text), Text1(1).Text) Then
                Cad = "       El fichero se ha generado satisfactoriamente.     " & vbCrLf & vbCrLf
                Cad = Cad & "     ¿  Desea imprimir el resultado de la remesa  ?" & vbCrLf & vbCrLf
                '--
                If MsgBox(Cad, vbDefaultButton1 + vbYesNoCancel) = vbYes Then
                    SQL = "numremes= " & Text1(0).Text & "|"
                    
                    frmImprimir.Opcion = 25
                    frmImprimir.NumeroParametros = 1
                    frmImprimir.FormulaSeleccion = SQL
                    frmImprimir.OtrosParametros = SQL
                    frmImprimir.SoloImprimir = False
                    frmImprimir.Show 'vbModal
                End If
    '        Else
    '            '-- La cosa ha ido malamente
    '            MsgBox "Se ha producido un error en la generación de la remesa, seguirá pendiente.", vbInformation
            End If
        Case 1 ' eliminacion de la remesa
             ' falta la comprobacion de errores
             
             SQL = "delete from sremes where numremes = " & Text1(0).Text
             Conn.Execute SQL
             
             SQL = "update sefect set numremes = null, fecremes = null, banremes = null "
             SQL = SQL & " where numremes = " & Text1(0).Text
             Conn.Execute SQL
             
        Case 2 ' abono de la remesa en contabilidad
            ' falta la contabilizacion dela bono en contabilidad
            If ContabilizarAbono(Text1(0).Text, Text1(1).Text) Then
                '-- La generación de remesas ha funcionado
                SQL = "UPDATE sremes SET situacio = 1 WHERE numremes= " & Text1(0).Text
                Conn.Execute SQL, , adCmdText
                MsgBox "La remesa ha sido abonada satisfactoriamente, su estado pasa a contabilizada", vbInformation
            End If
        End Select
    End If

End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(0).Text = Format(Text1(0).Text, "000000")
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
    Text2.Text = Format(Text2.Text, "dd/mm/yyyy")
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 0 Or Opcion = 2 Then Text1(1).Text = Format(Now, "dd/mm/yyyy")
    End If
        Screen.MousePointer = vbDefault
End Sub
'
Private Sub Form_Load()
Dim H As Single
Dim W As Single
Dim Acabar As Boolean

    Me.Top = 0
    Me.Left = 0

    Select Case Opcion
        Case 0
            Label1.Caption = "Diskette Remesa Norma 19"
            Label4(2).Caption = "F.Presentación"
            imgfec.Visible = True
            imgfec.Enabled = True
            Text1(1).Visible = True
            Text1(1).Enabled = True
            imgfec.Visible = True
            imgfec.Enabled = True
            
        Case 1
            Label1.Caption = "Eliminar Remesa"
            Label4(2).Caption = ""
            imgfec.Visible = False
            imgfec.Enabled = False
            Text1(1).Visible = False
            Text1(1).Enabled = False
            Label4(2).Caption = ""
            imgfec.Visible = False
            imgfec.Enabled = False
            
        Case 2
            Label1.Caption = "Abono Remesa"
            Label4(2).Caption = "Fecha de Abono"
            imgfec.Visible = True
            imgfec.Enabled = True
            Text1(1).Visible = True
            Text1(1).Enabled = True
            imgfec.Visible = True
            imgfec.Enabled = True
            
        End Select

    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    Limpiar Me
    W = Me.FrameRemesas.Width
    H = Me.FrameRemesas.Height
    
    Me.Width = W + 240
    Me.Height = H + 400

    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgfec_Click()
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(2).Text <> "" Then frmC.Fecha = CDate(Text1(2).Text)
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub ImgRem_Click(Index As Integer)
    
        Set frmM = New frmMensajes
        frmM.Opcion = 1
        frmM.Show
    
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)

End Sub


Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index))
    If Text1(Index) = "" Then Exit Sub
    
    Select Case Index
        Case 0
            If Text1(0).Text <> "" Then
                If EsNumerico(Text1(0).Text) Then
                        Text1(Index).Text = Format(Text1(Index).Text, "000000")
                        Text2.Text = DevuelveDesdeBD(1, "fecremes", "sremes", "numremes|", Text1(0).Text & "|", "N|", 1)
                End If
            End If
        Case 1
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta: " & Text1(Index), vbExclamation
                Text1(Index).Text = ""
                Text1(Index).SetFocus
            End If
            
     End Select
End Sub

'Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
'    ComprobarFechas = False
'    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
'        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Function
'        End If
'    End If
'    ComprobarFechas = True
'End Function

