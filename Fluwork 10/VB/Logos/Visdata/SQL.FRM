VERSION 5.00
Begin VB.Form frmSQL 
   Caption         =   "Instrucción SQL"
   ClientHeight    =   2880
   ClientLeft      =   3690
   ClientTop       =   1575
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016144
   Icon            =   "SQL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2863.353
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5147.588
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaveQueryDef 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   3375
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdExecuteSQL 
      Caption         =   "&Ejecutar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   15
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   30
      Width           =   1575
   End
   Begin VB.CommandButton cmdClearSQL 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   1695
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   30
      Width           =   1575
   End
   Begin VB.TextBox txtSQLStatement 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   435
      Width           =   5055
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Instrucción SQL"
Const BUTTON1 = "&Ejecutar"
Const BUTTON2 = "&Borrar"
Const BUTTON3 = "&Guardar"
Const MSG1 = "Actualizar"
Const MSG2 = "Escriba el nombre del QueryDef:"
Const MSG3 = "¿Es una QueryDef de SQLPassThrough?"
Const MSG4 = "Escriba el valor de la propiedad Connect:"
Const MSG5 = "¿La consulta está devolviendo filas?"
'>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub cmdClearSQL_Click()
  txtSQLStatement.Text = vbNullString
  txtSQLStatement.SetFocus
End Sub

Private Sub cmdSaveQueryDef_Click()
  On Error GoTo SQDErr

  Dim sQueryName As String
  Dim sTmp As String
  Dim qdNew As QueryDef

  If Not gnodDBNode Is Nothing Then
    If gnodDBNode.Tag = QUERY_STR Then
      'hay una Querydef seleccionada, de modo que el usuario quiere actualizar su consulta SQL
      If MsgBox(MSG1 & " '" & gnodDBNode.Text & "'?", vbYesNo + vbQuestion) = vbYes Then
        'almacena la consulta SQL de la ventana SQL en la
        'Querydef actualmente seleccionada
        gdbCurrentDB.QueryDefs(gnodDBNode.Text).SQL = Me.txtSQLStatement.Text
        Exit Sub
      End If
    End If
  End If
  
  'no hay Querydef seleccionada actualmente o el usuario
  'no quiso actualizar la actual, así que tenemos que
  'solicitar una nueva
  sQueryName = InputBox(MSG2)
  If Len(sQueryName) = 0 Then Exit Sub
  
  'comprueba si hay duplicados y sale si el usuario no va a sobrescribirlo
  If DupeTableName(sQueryName) Then
    Exit Sub
  End If
    
  'agrega la nueva Querydef
  Set qdNew = gdbCurrentDB.CreateQueryDef(sQueryName)
  'solicita una Querydef específica de SQL
  If MsgBox(MSG3, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
    sTmp = InputBox(MSG4)
    If Len(sTmp) > 0 Then
      qdNew.Connect = sTmp
      If MsgBox(MSG5, vbYesNo + vbQuestion) = vbNo Then
        qdNew.ReturnsRecords = False
      End If
    End If
  End If
  qdNew.SQL = txtSQLStatement.Text
  
  gdbCurrentDB.QueryDefs.Refresh
  RefreshTables Nothing

  Exit Sub

SQDErr:
  ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'sólo guarda el texto SQL si no contiene retornos de carro,
  'puesto que no se guardan en el archivo INI
'  If InStr(frmSQL.txtSQLStatement.Text, Chr(13)) = 0 Then
    SaveSetting APP_CATEGORY, App.Title, "SQLStatement", frmSQL.txtSQLStatement.Text
'  End If
  If frmSQL.WindowState = vbNormal Then
    SaveSetting APP_CATEGORY, App.Title, "SQLWindowTop", frmSQL.Top
    SaveSetting APP_CATEGORY, App.Title, "SQLWindowLeft", frmSQL.Left
    SaveSetting APP_CATEGORY, App.Title, "SQLWindowWidth", frmSQL.Width
    SaveSetting APP_CATEGORY, App.Title, "SQLWindowHeight", frmSQL.Height
  End If
End Sub

Private Sub txtSQLStatement_Change()
  cmdExecuteSQL.Enabled = Len(txtSQLStatement.Text) > 0
End Sub

Private Sub cmdExecuteSQL_Click()
  If Len(txtSQLStatement.Text) = 0 Then Exit Sub

  OpenQuery txtSQLStatement.Text, True
End Sub

Private Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdExecuteSQL.Caption = BUTTON1
  cmdClearSQL.Caption = BUTTON2
  cmdSaveQueryDef.Caption = BUTTON3
  
  txtSQLStatement.Text = GetINIString("SQLStatement", vbNullString)
 
  Me.Height = Val(GetINIString("SQLWindowHeight", "3000"))
  Me.Width = Val(GetINIString("SQLWindowWidth", "5370"))
  Me.Top = Val(GetINIString("SQLWindowTop", "0"))
  Me.Left = Val(GetINIString("SQLWindowLeft", "3850"))

End Sub

Private Sub Form_Resize()
  On Error Resume Next

  If WindowState <> vbMinimized Then
    txtSQLStatement.Width = Me.ScaleWidth - (txtSQLStatement.Left * 2)
    txtSQLStatement.Height = Me.ScaleHeight - (txtSQLStatement.Top + 50)
  End If
End Sub

Private Sub txtSQLStatement_DragDrop(Source As Control, x As Single, y As Single)
  If Source = frmDatabase.tvDatabase Then
    frmSQL.txtSQLStatement.Text = gdbCurrentDB.QueryDefs(gnodDBNode.Text).SQL
  End If
End Sub
