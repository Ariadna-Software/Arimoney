VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoASegurado 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Aviso siniestro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Aviso falta de pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   3735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2135
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmListadoASegurado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.- Los dos
    ' 1.- Solo Falta pago
    ' 2.- solo Siniestro

Dim PrimVez As Boolean
Dim SQL As String


Private Sub Command1_Click()

    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        CargaListviewAsegurados
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimVez = True
    Me.Icon = frmPpal.Icon
    
    Me.Caption = "Listado falta aviso asegurado   -   " & Format(Now, "dd/mm/yyyy")
    
    Me.Option1(0).Visible = Opcion <> 2
    Me.Option1(1).Visible = Opcion <> 1
    If Opcion <> 2 Then Me.Option1(0).Value = True
    
    Set Me.ListView1.SmallIcons = frmPpal.ImgListviews
End Sub

Private Sub CargaListviewAsegurados()
Dim I As Integer
Dim IT
Dim Dias As Integer
Dim F As Date

    Set miRsAux = New ADODB.Recordset


    CargaColumnas

    'Monta EL SQL
    MontaSQLAvisosSeguros Option1(0).Value, SQL
    
    SQL = "Select numserie,codfaccl,fecfaccl,fecvenci, scobro.codmacta ,nommacta,impvenci,numorden,devuelto,fecprorroga " & SQL
    SQL = SQL & " ORDER  BY fecvenci ,impvenci"
    
    Me.ListView1.ListItems.Clear
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!codfaccl, "000000")
        IT.SubItems(2) = miRsAux!fecfaccl
        IT.SubItems(3) = miRsAux!fecvenci
        
        
        
        
        
        
        I = 5
        If Option1(0).Value Then
            F = DateAdd("d", vParam.DiasMaxAvisoH, miRsAux!fecvenci)  'teneia que avisar F dia
            Dias = DateDiff("d", Now, F)  'hace dias
        
        
        Else
            'SINIESTRO
            If IsNull(miRsAux!fecprorroga) Then
                IT.SubItems(I) = " "
                F = DateAdd("d", vParam.DiasMaxSiniestroH, miRsAux!fecvenci)   'teneia que avisar F dia
            Else
                IT.SubItems(I) = miRsAux!fecprorroga
                F = DateAdd("d", vParam.DiasAvisoDesdeProrroga, miRsAux!fecprorroga)    'teneia que avisar F dia
            End If
            I = 6
            
            Dias = DateDiff("d", Now, F)  'hace dias
            
        End If
        IT.SubItems(4) = Dias
        IT.SubItems(I) = miRsAux!codmacta
        IT.SubItems(I + 1) = miRsAux!Nommacta
        IT.SubItems(I + 2) = Format(miRsAux!impvenci, FormatoImporte)
        IT.SubItems(I + 3) = miRsAux!numorden
        
        If Val(miRsAux!devuelto) = 1 Then
            IT.SmallIcon = 1
        Else
            IT.SmallIcon = 2
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    

End Sub

Private Sub ListView1_DblClick()



    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    With Me.ListView1.SelectedItem
        
        SQL = "numserie='" & .Text & "' AND codfaccl = " & .SubItems(1) & " AND fecfaccl='" & Format(.SubItems(2), FormatoFecha) & "' AND numorden ="
        
        If Me.Option1(1).Value Then
            frmAseguradosAccion.Opcion = 1
            frmAseguradosAccion.Label1.Caption = .SubItems(7)
            frmAseguradosAccion.lblTitulo = Me.Option1(1).Caption
            SQL = SQL & .SubItems(9)
        Else
            frmAseguradosAccion.Opcion = 0
            frmAseguradosAccion.Label1.Caption = .SubItems(6)
            frmAseguradosAccion.lblTitulo = Me.Option1(0).Caption
            SQL = SQL & .SubItems(8)
        End If
        frmAseguradosAccion.Label2.Caption = Trim(.Text & .SubItems(1)) & "  de " & .SubItems(2) & "    Vto: " & .SubItems(3)
        
        frmAseguradosAccion.SQLVto = SQL
        frmAseguradosAccion.Show vbModal
        If CadenaDesdeOtroForm <> "" Then Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
        
    End With
End Sub

Private Sub Option1_Click(Index As Integer)
    
    CargaListviewAsegurados
End Sub



Private Sub CargaColumnas()
Dim colX As ColumnHeader
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim NCols As Integer
Dim I As Integer

        ListView1.ColumnHeaders.Clear
   
        If Me.Option1(0).Value Then
            NCols = 9
            Columnas = "Serie|Factura|F. Factura|F. VTO|Dias|Cuenta|Nombre|Importe|NumVenci|"
            Ancho = "800|1000|1100|1100|600|1200|100%|1000|0|"
            ALIGN = "LLLLDLLDD"
        
        Else
            NCols = 10
            Columnas = "Serie|Factura|F. Factura|F. VTO|Dias|F.Prorroga|Cuenta|Nombre|Importe|NumVenci|"
            Ancho = "700|900|1100|1100|600|1100|1200|90%|1000|0|"
            ALIGN = "LLLLLDLLDD"
        
        End If
        ListView1.Tag = 7600    'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
        
        
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
         SQL = RecuperaValor(Columnas, I)
         If SQL <> "" Then
             Set colX = ListView1.ColumnHeaders.Add()
             colX.Text = SQL
             'ANCHO
             SQL = RecuperaValor(Ancho, I)
             colX.Tag = SQL
             'align
             SQL = Mid(ALIGN, I, 1)
             If SQL = "L" Then
                 'NADA. Es valor x defecto
             Else
                 If SQL = "D" Then
                     colX.Alignment = lvwColumnRight
                 Else
                     'CENTER
                     colX.Alignment = lvwColumnCenter
                 End If
             End If
         End If
     Next I



    
    ListView1.Tag = ListView1.Width - ListView1.Tag - 320 'Del margen
    For I = 1 To Me.ListView1.ColumnHeaders.Count
        If InStr(1, ListView1.ColumnHeaders(I).Tag, "%") Then
            SQL = (Val(ListView1.ColumnHeaders(I).Tag) * (Val(ListView1.Tag)) / 100)
        Else
            'Si no es de % es valor fijo
            SQL = Val(ListView1.ColumnHeaders(I).Tag)
        End If
        Me.ListView1.ColumnHeaders(I).Width = Val(SQL)
    Next I



End Sub
