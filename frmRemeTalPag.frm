VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemeTalPag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remesas talón-pagaré"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView2 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero Ref."
         Object.Width           =   5293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Banco"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "F. recepcion"
         Object.Width           =   2029
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "F. Vto"
         Object.Width           =   2029
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cta"
         Object.Width           =   1854
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cliente"
         Object.Width           =   3149
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Importe"
         Object.Width           =   2207
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   7080
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   11400
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   10080
      TabIndex        =   3
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   12735
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   1
         Left            =   7200
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   270
         Width           =   5535
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   10560
         Top             =   315
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   10200
         Picture         =   "frmRemeTalPag.frx":0000
         ToolTipText     =   "Editar referencia del talon/pagare"
         Top             =   315
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   8880
         Picture         =   "frmRemeTalPag.frx":0A02
         ToolTipText     =   "Quitar seleccion"
         Top             =   315
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   9360
         Picture         =   "frmRemeTalPag.frx":0B4C
         ToolTipText     =   "Seleccionar todos"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   303
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   4895
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Serie"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factura"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "F. Fact"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Vto"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "F. Vto"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cuenta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cliente"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Importe"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Num tal"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Vencimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Label Label5 
      Caption         =   "Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   7080
      Width           =   1695
   End
End
Attribute VB_Name = "frmRemeTalPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vRemesa As String 'Si viene="" NUEVA
                         'si no indica remesa,año
Public SQL As String
Public Talon As Boolean  'talon o pagare
Dim Importe As Currency
Dim CodRem As Integer
Dim AnyoRem As Integer


Private Sub Command1_Click(Index As Integer)
Dim YaRemesado As Currency
Dim Limite As Currency
Dim I As Integer
    If Index = 0 Then
        'Descrip`cion obligada
        If Text2.Text = "" Then
            MsgBox "El campo descripcion debe tener valor", vbExclamation
            Exit Sub
        End If
    
        'Crear remesa talon pagare
        Importe = 0
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Checked Then
                'Este documento. Vemos el importe del documento
                'Antes septiembre 2009
                'SQL = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", ListView2.ListItems(NumRegElim).Text)
                'If SQL = "" Then SQL = "0"   'No deberia pasar
                SQL = ListView2.ListItems(NumRegElim).SubItems(7)
                Importe = Importe + CCur(SQL)
            End If
        Next

                
        
        
        If Importe = 0 Then
            MsgBox "Seleccione algun talón/pagaré", vbExclamation
            Exit Sub
        End If
        
        'La fecha y las cuentas bloqueadas ya las hemos comprobado en la fase anterior
        'Ahora el limite del banco
        If Talon Then
            SQL = "talonriesgo"
            NumRegElim = 3 '   para la select de abajo
        Else
            NumRegElim = 2
            SQL = "pagareriesgo" 'para la select de abajo
        End If
        SQL = DevuelveDesdeBD(SQL, "ctabancaria", "codmacta", Trim(Mid(Text1(0).Text, 1, 10)), "T")
        If SQL <> "" Then
            Limite = CCur(SQL)
        Else
            Limite = -1
        End If
        
        'Tenemos que ver todos los vencimientos que sean de tipo de pago talon o pagare, que la cta de pago sea
        'la del banco en question y ver cuanto llevamos remesado
        SQL = "select sum(impcobro) FROM scobro,sforpa WHERE scobro.codforpa = sforpa.codforpa AND "
        SQL = SQL & "siturem>'B' AND siturem < 'Z'"

        SQL = SQL & " and ctabanc2='" & Trim(Mid(Text1(0).Text, 1, 10)) & "' AND tiporem = " & NumRegElim
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        YaRemesado = 0
        If Not miRsAux.EOF Then
            'Le sumo lo que llevamos en esta remesa (los k estan check) a los vtos ya remesados Y nO eleminidados
            YaRemesado = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        
        If Limite >= 0 Then
            If Limite - (Importe + YaRemesado) < 0 Then
                'Supera el riesgo
                SQL = "Esta superando el riesgo permitido por el banco." & vbCrLf
                SQL = SQL & "Riesgo concedido: " & Format(Limite, FormatoImporte) & vbCrLf
                SQL = SQL & "Remesa: " & Format(Importe, FormatoImporte) & vbCrLf
                SQL = SQL & "Ya remesado: " & Format(YaRemesado, FormatoImporte) & vbCrLf
                
                SQL = SQL & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
        End If
        
        Set miRsAux = New ADODB.Recordset
        
        'UNa ultima comprobacion. Vamos a ver si un mismo vencimiento esta en dos docuemntos
        'distintos, o si alguno de los vencimientos pertecence a una remesa que aun no ha sido
        'borrada
        If Not ComprobarEfectosCobradosParcialmente Then
            Set miRsAux = New ADODB.Recordset
            Exit Sub
        End If
        
        
        
        
        Conn.BeginTrans
        'Remesamos ---------------------------------------
        If Not EfectuarRemesa Then
            Conn.RollbackTrans
            Exit Sub
        Else
            Conn.CommitTrans
        End If
        Set miRsAux = Nothing
    Else
        'Cancelar. Si ha cambiado algo le pregnto
        If Me.Tag <> "" Then
            SQL = "Ha efectuado cambios. Descartar los cambios?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub

        
        
        
        
        'UNa ultima comprobacion. Vamos a ver si un mismo vencimiento esta en dos docuemntos
        'distintos, o si alguno de los vencimientos pertecence a una remesa que aun no ha sido
        'borrada
Private Function ComprobarEfectosCobradosParcialmente() As Boolean
Dim AUX As String
Dim MasDeUnDocumento As Byte
    On Error GoTo EComprobarEfectosCobradosParcialmente
    ComprobarEfectosCobradosParcialmente = False
    
    
        AUX = ""
        MasDeUnDocumento = 0
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Checked Then
                'Este documento. Vemos el importe del documento
                AUX = AUX & "," & ListView2.ListItems(NumRegElim).Text
                If MasDeUnDocumento = 0 Then
                    MasDeUnDocumento = 1
                Else
                    MasDeUnDocumento = 2
                End If
            End If
        Next
        
        AUX = Mid(AUX, 2) 'quito la primera coma
        If MasDeUnDocumento > 1 Then
            
            '1. Si existe algun vto cobrado parcialmente y recepcionado en dos de los documentos que vamos a recepcionar
            SQL = "Select scobro.numserie,scobro.codfaccl,scobro.fecfaccl,scobro.numorden,count(*)"
            SQL = SQL & " FROM slirecepdoc left join scobro on scobro.numserie=slirecepdoc.numserie AND codfaccl=numfaccl and"
            SQL = SQL & " scobro.fecfaccl = slirecepdoc.fecfaccl And numorden = numvenci"
            SQL = SQL & " WHERE id in (" & AUX & ") group by 1,2,3,4 having count(*) >1"
        
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                SQL = SQL & miRsAux!NUmSerie & miRsAux!codfaccl & " / " & miRsAux!numorden & vbCrLf
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
            If SQL <> "" Then
                SQL = "Los siguientes vencimientos estan mas de una vez: " & vbCrLf & SQL & vbCrLf
                SQL = SQL & "No deberia seguir con el proceso. ¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
                'Exit Sub  'FALTA### ver si hay que salir
            End If
        End If
        
        'Veremos si los vtos estan ya remesados
        SQL = "Select scobro.numserie,scobro.codfaccl,scobro.fecfaccl,scobro.numorden"
        SQL = SQL & " FROM slirecepdoc left join scobro on scobro.numserie=slirecepdoc.numserie AND codfaccl=numfaccl and"
        SQL = SQL & " scobro.fecfaccl = slirecepdoc.fecfaccl And numorden = numvenci and codrem>0"
        SQL = SQL & " WHERE id in (" & AUX & ") group by 1,2,3,4"
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            If Not IsNull(miRsAux!NUmSerie) Then SQL = SQL & miRsAux!NUmSerie & miRsAux!codfaccl & " / " & miRsAux!numorden & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
        If SQL <> "" Then
            SQL = "Los siguientes vencimientos estan remesados y no ha sido eliminado el riesgo: " & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Exit Function
        End If
        
    ComprobarEfectosCobradosParcialmente = True
    
    Exit Function
EComprobarEfectosCobradosParcialmente:
    MuestraError Err.Number, Err.Description
End Function

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    Me.Tag = ""   'Si no es '' entonces ha cambiado algo
    Text2.Text = ""
    CargaIconoListview Me.ListView1
    CargaIconoListview Me.ListView2
    Carga1ImagenAyuda Image1(1), 3 'ayuda
    Set miRsAux = New ADODB.Recordset
    CargaDatos
    Set miRsAux = Nothing
End Sub






Private Sub CargaDatos()
    Dim IT As ListItem
    
    ListView1.ListItems.Clear
    SQL = Replace(SQL, "id", "codigo")
    SQL = "Select * from scarecepdoc,cuentas where scarecepdoc.codmacta = cuentas.codmacta AND " & SQL
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView2.ListItems.Add()
        IT.Text = miRsAux!Codigo
        IT.SubItems(1) = miRsAux!numeroref
        IT.SubItems(2) = DBLet(miRsAux!banco, "T") & " "
        IT.SubItems(3) = Format(miRsAux!fecharec, "dd/mm/yyyy")
        IT.SubItems(4) = Format(miRsAux!fechavto, "dd/mm/yyyy")
        IT.SubItems(5) = miRsAux!codmacta
        IT.SubItems(6) = miRsAux!Nommacta
        IT.SubItems(7) = Format(miRsAux!Importe, FormatoImporte)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If ListView2.ListItems.Count > 0 Then
        Set ListView2.SelectedItem = ListView2.ListItems(1)
        CargaDatosLineas
    End If
End Sub


Private Sub CargaDatosLineas()
Dim IT As ListItem
    On Error GoTo EC
    
    If vRemesa <> "" Then
        Text2.Text = RecuperaValor(vRemesa, 3)
        CodRem = Val(RecuperaValor(vRemesa, 1))
        AnyoRem = Val(RecuperaValor(vRemesa, 2))
    Else
        CodRem = 0
    End If
    
    ListView1.ListItems.Clear
    SQL = "Select scobro.numserie,codfaccl,scobro.fecfaccl,fecvenci, numorden,impvenci ,gastos ,impcobro,reftalonpag,codrem,anyorem  "
    SQL = SQL & " FROM slirecepdoc left join scobro on scobro.numserie=slirecepdoc.numserie AND codfaccl=numfaccl and"
    SQL = SQL & " scobro.fecfaccl=slirecepdoc.fecfaccl AND numorden=numvenci"
    SQL = SQL & " WHERE id= " & ListView2.SelectedItem.Text
        
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        If IsNull(miRsAux!NUmSerie) Then
            'ERROR GRAVE. Hay un vto en las lineas del docuemnto que NO esta en
            IT.ForeColor = vbRed
            IT.Bold = True
            IT.Text = "ERR"
            For NumRegElim = 1 To ListView1.ColumnHeaders.Count - 1
                IT.SubItems(NumRegElim) = "ERROR"
                IT.ListSubItems(NumRegElim).ForeColor = vbRed
                
                IT.ListSubItems(NumRegElim).Bold = True
            Next NumRegElim
        Else
             IT.Text = Mid(DBLet(miRsAux!NUmSerie, "T") & "   ", 1, 3)
             IT.SubItems(1) = Format(miRsAux!codfaccl, "000000")
             IT.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
             IT.SubItems(3) = miRsAux!numorden
             IT.SubItems(4) = Format(miRsAux!fecvenci, "dd/mm/yyyy")
            ' IT.SubItems(5) = miRsAux!codmacta
            ' IT.SubItems(6) = miRsAux!Nommacta
             'Lo debe cojer de impcobro
             IT.SubItems(7) = Format(miRsAux!impcobro, FormatoImporte)
             
             IT.SubItems(8) = DBLet(miRsAux!reftalonpag, "T")
             
             If CodRem > 0 Then
                 If Not IsNull(miRsAux!CodRem) Then
                     If Val(miRsAux!CodRem) = CodRem And Val(miRsAux!AnyoRem) = AnyoRem Then
                         'Voy a pintar de colorines el vto
                         IT.ForeColor = vbRed
                         For NumRegElim = 1 To IT.ListSubItems.Count
                             IT.ListSubItems(NumRegElim).ForeColor = vbRed
                         Next NumRegElim
                         IT.Checked = True
                     End If
                 End If
             End If
         End If 'de null numserie
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set ListView1.SelectedItem = Nothing
    PonerfocoObj Me.Command1(1)
    Exit Sub
EC:
    MuestraError Err.Number, "Carga datos"
End Sub

'Private Sub Image1_Click(Index As Integer)
'    If Index = 1 Then
'        SQL = "Editar referencia talón/pagaré" & vbCrLf & vbCrLf
'        SQL = SQL & "Permitirá añadir una referencia al talón-pagaré(numero, fecha...)" & vbCrLf
'        SQL = SQL & "Cuando acepte la referencia, le pedirá la siguiente." & vbCrLf
'        SQL = SQL & "Para eliminar la referencia ya asignada, edítela y ponga un espacio en blanco." & vbCrLf
'        SQL = SQL & "Con la opcion 'cancelar' termina de introducir" & vbCrLf
'        MsgBox SQL, vbInformation
'        Exit Sub
'    End If
'
'
'    If ListView1.SelectedItem Is Nothing Then Exit Sub
'
'    EditarNodo ListView1.SelectedItem.Index
'
'End Sub



'Voy a modificar esta variable
'Private Sub EditarNodo(ByVal PrimerNodo As Integer)
'Dim C As String
'
'    Exit Sub   'Asi NUNCA entra aqui
'
'    Do
'        With ListView1.ListItems(PrimerNodo)
'            SQL = "Cliente: " & .SubItems(5) & " - " & .SubItems(6) & vbCrLf
'            SQL = SQL & "Vencimiento: " & .Text & .SubItems(1) & " / " & .SubItems(3) & vbCrLf
'            SQL = SQL & "Fecha: " & .SubItems(2) & vbCrLf
'            SQL = SQL & "Importe: " & .SubItems(7) & vbCrLf
'            'La referencia(si es k tiene)
'            C = .SubItems(8)
'
'
'            C = InputBox(SQL, "Ref. talon/pagare)", C)
'            If C <> "" Then
'                C = Trim(C)
'                .SubItems(8) = C
'                .EnsureVisible
'                'Siguiente
'                PrimerNodo = PrimerNodo + 1
'                'Nos salimos
'                If PrimerNodo > ListView1.ListItems.Count Then PrimerNodo = 0
'                Me.Tag = "C" 'Ha cambiado cosas
'            Else
'                'Ha cancelado
'                PrimerNodo = 0 'Para salir
'            End If
'        End With
'    Loop Until PrimerNodo = 0
'End Sub



Private Sub imgCheck_Click(Index As Integer)
    For NumRegElim = 1 To ListView1.ListItems.Count
        ListView1.ListItems(NumRegElim).Checked = Index = 1
    Next
    If Index = 1 Then Me.Tag = "C"
End Sub

'Private Sub ListView1_DblClick()
'    If ListView1.SelectedItem Is Nothing Then Exit Sub
'    EditarNodo ListView1.SelectedItem.Index
'End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Me.Tag = "C" 'Para saber que ha cambiado
End Sub


Private Function EfectuarRemesa() As Boolean
Dim TipoRemesa As Byte
Dim R As ADODB.Recordset
    On Error GoTo EEfectuarRemesa
    EfectuarRemesa = False
    '---------------------------------------------------
    'Creamos la remesa
    SQL = "Select nomcta as numeroremesa,cta from tmpCierre1 where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Datos remesa NO encontrados.", vbExclamation
        miRsAux.Close
        Exit Function
    End If
    If vRemesa = "" Then
        CodRem = miRsAux!NumeroRemesa
        AnyoRem = Year(CDate(Text1(1).Text))
    End If
    If Talon Then
        TipoRemesa = 3
    Else
        TipoRemesa = 2
    End If

    
    'Si estamos modificando la remesa tenemos que quitar la marca de remeados
    If vRemesa <> "" Then
        SQL = "UPDATE  scobro SET siturem= NULL,codrem= NULL, anyorem =NULL,tiporem = NULL"
        SQL = SQL & " WHERE codrem = " & CodRem & " and anyorem =" & AnyoRem
        Conn.Execute SQL
    End If

    Set R = New ADODB.Recordset
    
    'Updateamos los vencimientos.  Desde el listview2 vemos que documentos esta llevando al banco
    For NumRegElim = 1 To ListView2.ListItems.Count
        
            If ListView2.ListItems(NumRegElim).Checked Then
                SQL = "Select scobro.numserie,codfaccl,scobro.fecfaccl,fecvenci, numorden,impvenci ,gastos ,impcobro,reftalonpag,codrem,anyorem  "
                SQL = SQL & " FROM slirecepdoc left join scobro on scobro.numserie=slirecepdoc.numserie AND codfaccl=numfaccl and"
                SQL = SQL & " scobro.fecfaccl=slirecepdoc.fecfaccl AND numorden=numvenci"
                SQL = SQL & " WHERE id= " & ListView2.ListItems(NumRegElim).Text
    
                R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not R.EOF
                    SQL = R!NUmSerie 'para que de el error si no existe
                    
                    
                    'La situacion entra directamente a cancelacion cliente
                    SQL = "UPDATE  scobro SET siturem= 'F',codrem= " & CodRem & ", anyorem =" & AnyoRem & ","
                    SQL = SQL & " tiporem = " & TipoRemesa

                    'ponemos la cuenta de banco donde va remesado
                    SQL = SQL & ", ctabanc2 ='" & miRsAux!Cta & "' "
                    'Por si acaso a puesto talon referencia
                    SQL = SQL & " WHERE numserie = '" & R!NUmSerie & "' and codfaccl = "
                    SQL = SQL & R!codfaccl & " and fecfaccl ='" & Format(R!fecfaccl, FormatoFecha)
                    SQL = SQL & "' AND numorden =" & R!numorden
                                
                    Conn.Execute SQL
                    R.MoveNext
                Wend
                R.Close
            End If

    Next NumRegElim

    'Cremos la cabcera de las remesas
    If vRemesa = "" Then
        SQL = "insert into `remesas` (`codigo`,`anyo`,`fecremesa`,`fecini`,`fecfin`,`situacion`,`codmacta`,`tipo`,`importe`,`descripcion`,`tiporem`) values ("
        SQL = SQL & miRsAux!NumeroRemesa & "," & Year(CDate(Text1(1).Text)) & ",'" & Format(Text1(1).Text, FormatoFecha) & "',NULL,NULL,'F','"
        SQL = SQL & miRsAux.Fields!Cta & "',NULL," & TransformaComasPuntos(CStr(Importe)) & ",'" & DevNombreSQL(Text2.Text) & "'," & TipoRemesa & ")"
    Else
        'Updatemaos
        SQL = "UPDATE remesas set importe=" & TransformaComasPuntos(CStr(Importe))
        SQL = SQL & ", descripcion = '" & DevNombreSQL(Text2.Text) & "'"
        SQL = SQL & " WHERE codigo = " & CodRem & " AND anyo = " & AnyoRem
    End If
    Conn.Execute SQL

    'Marco en scarecepdoc el llevada a banco
     For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Checked Then
            
                SQL = "UPDATE scarecepdoc SET  LlevadoBanco = 1 WHERE codigo = " & ListView2.ListItems(NumRegElim).Text
                Conn.Execute SQL
            End If

    Next NumRegElim
    miRsAux.Close
    
    EfectuarRemesa = True
    Set R = Nothing
    Exit Function
EEfectuarRemesa:
    MuestraError Err.Number, Err.Description
    Set R = Nothing
End Function





Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set ListView2.SelectedItem = Item
    Set miRsAux = New ADODB.Recordset
    CargaDatosLineas
    Set miRsAux = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub
