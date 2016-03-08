VERSION 5.00
Begin VB.Form aridoc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "aridoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    '   insertarMuchos
    
End Sub


Private Sub insertarMuchos()
Dim SQL As String
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Long
Dim J As Integer
    SQL = "INSERT INTO Aridoc.tmp1 (codigo, campo1,fecha1,campo3,extension, campo2,fecha2,    observaciones,  importeD, importeH, Grupo, otros) VALUES ("
    
    For i = 1 To 3000000
        cad = i & ",'Campo " & i & "','" & Format(Now, FormatoFecha) & "',NULL,"
        J = i Mod 7
        cad = cad & J & ","
        If J = 0 Then
            cad = cad & "'Segun: " & i & " de otro " & i & "','" & Format(Now, FormatoFecha) & "',"
            cad = cad & "'OBSERAVIONES varias para " & i & "'"
        Else
            cad = cad & "NULL,NULL,NULL"
        End If
        'Importe 2
        If J < 4 Then
            cad = cad & ",1000,NULL,"
        Else
            cad = cad & ",NULL,1320.2,"
        End If
        
        'Grupos y otros
        J = i Mod 32
        Caption = i & " de 3000000"
        If J = 0 Then Me.Show
        'grupo
        cad = cad & J & ","
        'otros
        J = J \ 2
        cad = cad & J & ")"
        cad = SQL & cad
        Conn.Execute cad
        
    Next i
        
    Exit Sub
End Sub
