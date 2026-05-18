VERSION 5.00
Begin VB.Form frmQuerys 
   Caption         =   "Consultas"
   ClientHeight    =   4185
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5100
   Tag             =   "Querys"
   Begin VB.ListBox lstQueryDefs 
      Height          =   1260
      Left            =   96
      TabIndex        =   0
      Top             =   274
      Width           =   3392
   End
   Begin VB.TextBox txtSQLStatement 
      BackColor       =   &H00FFFFFF&
      Height          =   2159
      Left            =   96
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1921
      Width           =   4931
   End
   Begin VB.CommandButton cmdRemoveQuery 
      Caption         =   "Eli&minar"
      Height          =   370
      Left            =   3572
      TabIndex        =   3
      Tag             =   "&Remove"
      Top             =   1277
      Width           =   1443
   End
   Begin VB.CommandButton cmdSaveQueryDef 
      Caption         =   "&Guardar"
      Height          =   370
      Left            =   3572
      TabIndex        =   2
      Tag             =   "&Save"
      Top             =   775
      Width           =   1443
   End
   Begin VB.CommandButton cmdExecuteSQL 
      Caption         =   "&Ejecutar"
      Enabled         =   0   'False
      Height          =   370
      Left            =   3572
      TabIndex        =   1
      Tag             =   "&Execute"
      Top             =   274
      Width           =   1443
   End
   Begin VB.Label lblSQL 
      Caption         =   "Instrucción SQL:"
      Height          =   251
      Index           =   1
      Left            =   132
      TabIndex        =   6
      Tag             =   "SQL Statement:"
      Top             =   1682
      Width           =   2189
   End
   Begin VB.Label lblSQL 
      Caption         =   "Consultas almacenadas:"
      Height          =   255
      Index           =   0
      Left            =   96
      TabIndex        =   5
      Tag             =   "Saved Querys:"
      Top             =   45
      Width           =   2190
   End
End
Attribute VB_Name = "frmQuerys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================================
'esta plantilla necesita el siguiente código (o su equivalente)
'para estar presente en la aplicación, así como una referencia a DAO 3.50
'y la plantilla Cuadrícula de datos
'
'Public gsDatabase As String
'Public gsRecordsource As String
'
'Sub Main()
'  gsDatabase = "c:\vb5\biblio.mdb"
'  frmQuerys.Show
'End Sub
'====================================================================


Dim mdbDatabase As Database
Private Sub Form_Load()
    Set mdbDatabase = OpenDatabase(gsDatabase)
    RefreshQuerys
    Me.Left = GetSetting(App.Title, "Settings", "QueryLeft", 0)
    Me.Top = GetSetting(App.Title, "Settings", "QueryTop", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "QueryLeft", Me.Left
        SaveSetting App.Title, "Settings", "QueryTop", Me.Top
    End If
End Sub



Private Sub cmdSaveQueryDef_Click()
    On Error GoTo SQDErr

    Dim sQueryName As String
    Dim sTmp As String
    Dim qdNew As QueryDef

    If lstQueryDefs.ListIndex >= 0 Then
        'se selecciona un querydef para que el usuario pueda actualizar su SQL
        If MsgBox("¿Desea actualizar '" & lstQueryDefs.Text & "'?", vbYesNo + vbQuestion) = vbYes Then
            'almacena la instrucción SQL desde la ventana de SQL al
            'querydef seleccionado actualmente
            mdbDatabase.QueryDefs(lstQueryDefs.Text).SQL = Me.txtSQLStatement.Text
            Exit Sub
        End If
    End If
    
    'tanto si no hay un querydef seleccionado actualmente por el usuario
    'como si no se desea actualizar el actual es necesario
    'solicitar un nuevo nombre
    sQueryName = InputBox("Escriba el nuevo nombre de la consulta:")
    If Len(sQueryName) = 0 Then Exit Sub
    
    'agrega el nuevo querydef
    Set qdNew = mdbDatabase.CreateQueryDef(sQueryName)
    'pregunta si es un querydef de tipo passthrough
    If MsgBox("¿Es un QueryDef de tipo SQLPassThrough?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        sTmp = InputBox("Escriba el nombre de la propiedad Connect:")
        If Len(sTmp) > 0 Then
            qdNew.Connect = sTmp
            If MsgBox("¿La consulta devuelve valores?", vbYesNo + vbQuestion) = vbNo Then
                qdNew.ReturnsRecords = False
            End If
        End If
    End If
    qdNew.SQL = txtSQLStatement.Text
    
    mdbDatabase.QueryDefs.Refresh
    RefreshQuerys

    Exit Sub

SQDErr:
    MsgBox Err.Description
End Sub

Private Sub lstQueryDefs_Click()
    txtSQLStatement.Text = mdbDatabase.QueryDefs(lstQueryDefs.Text).SQL
End Sub

Private Sub lstQueryDefs_DblClick()
    cmdExecuteSQL_Click
End Sub

Private Sub txtSQLStatement_Change()
    If Len(txtSQLStatement.Text) > 0 Then
        cmdExecuteSQL.Enabled = True
    Else
        cmdExecuteSQL.Enabled = False
    End If
End Sub

Private Sub cmdExecuteSQL_Click()
    Dim rsTmp As Recordset
    Dim dbTmp As Database
    Dim qdfTmp As QueryDef
    Dim bSavedQDF As Boolean
    Dim sSQL As String
    
    If Len(txtSQLStatement.Text) = 0 Then Exit Sub
    
    Set dbTmp = OpenDatabase(gsDatabase)
    
    If lstQueryDefs.ListIndex >= 0 Then
        sSQL = dbTmp.QueryDefs(lstQueryDefs.Text).SQL
        If sSQL = txtSQLStatement.Text Then
            Set qdfTmp = dbTmp.QueryDefs(lstQueryDefs.Text)
            bSavedQDF = True
            If Not SetQryParams(qdfTmp) Then Exit Sub
        Else
            'crea un querydef temporal
            Set qdfTmp = dbTmp.CreateQueryDef(vbNullString, txtSQLStatement.Text)
        End If
    Else
        'crea un querydef temporal
        Set qdfTmp = dbTmp.CreateQueryDef(vbNullString, txtSQLStatement.Text)
    End If

    Screen.MousePointer = vbHourglass
    If UCase(Mid(txtSQLStatement, 1, 6)) = "SELECT" And InStr(UCase(txtSQLStatement.Text), " INTO ") = 0 Then
        On Error GoTo SQLErr
MakeDynaset:
        Dim f As New frmDataGrid
        Set rsTmp = qdfTmp.OpenRecordset()
        Set f.Data1.Recordset = rsTmp
        If bSavedQDF Then
            f.Caption = qdfTmp.Name
        Else
            f.Caption = Left(txtSQLStatement.Text, 32) & "..."
        End If
        f.Show
    Else
        On Error GoTo SQLErr
        qdfTmp.Execute
    End If

    Screen.MousePointer = vbDefault
    Exit Sub

SQLErr:
    If Err = 3065 Or Err = 3078 Then 'devuelve filas o no se encontró el nombre, se intenta crear un recordset
        Resume MakeDynaset
    End If
    MsgBox Err.Description

SQLEnd:

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If WindowState <> 1 Then
        If Me.Width < 5220 Then Me.Width = 5220
        If Me.Height < 2784 Then Me.Height = 2784
        
        txtSQLStatement.Width = Me.Width - 320
        txtSQLStatement.Height = Me.Height - 2424
    End If
End Sub

Sub RefreshQuerys()
    Dim qdf As QueryDef
    
    lstQueryDefs.Clear
    
    For Each qdf In mdbDatabase.QueryDefs
        lstQueryDefs.AddItem qdf.Name
    Next
    
End Sub

Private Function SetQryParams(rqdf As QueryDef) As Boolean
    On Error GoTo SPErr
    
    Dim prm As Parameter
    Dim sTmp As String
    Dim i As Integer
    
    For Each prm In rqdf.Parameters
        'obtiene el valor del usuario
        sTmp = InputBox("Escriba un valor para el parámetro '" & prm.Name & "':")
        If Len(sTmp) = 0 Then
            'sale si el usuario no escribe alguno de los parámetros
            SetQryParams = False
            Exit Function
        End If
        'store the value
        prm.Value = CVar(sTmp)
    Next
    
    SetQryParams = True
    Exit Function
        
SPErr:
    MsgBox Err.Description
End Function

