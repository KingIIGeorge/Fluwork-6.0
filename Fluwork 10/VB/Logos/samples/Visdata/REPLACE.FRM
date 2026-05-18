VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sustitución global"
   ClientHeight    =   3855
   ClientLeft      =   1860
   ClientTop       =   2070
   ClientWidth     =   5175
   HelpContextID   =   2016091
   Icon            =   "REPLACE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3421.786
   ScaleMode       =   0  'User
   ScaleWidth      =   5081.244
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstTables 
      BackColor       =   &H00FFFFFF&
      Height          =   1650
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.ListBox lstFields 
      BackColor       =   &H00FFFFFF&
      Height          =   1650
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtReplaceWith 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox txtCondition 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   4695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Tabla:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Criterio:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1485
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Reemplazar con:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1965
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Campo:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   930
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Sustitución global"
Const BUTTON1 = "&Aceptar"
Const BUTTON2 = "&Cancelar"
Const Label1 = "&Tabla:"
Const Label2 = "Ca&mpo:"
Const LABEL3 = "&Reemplazar con:"
Const LABEL4 = "Cr&iterio:"
Const MSG1 = "Reemplazando registros"
'>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub lstFields_Click()
  cmdOK.Enabled = Len(lstFields.Text) > 0 And Len(txtReplaceWith.Text) > 0
End Sub

Private Sub txtReplaceWith_Change()
  cmdOK.Enabled = Len(lstFields.Text) > 0 And Len(txtReplaceWith.Text) > 0
End Sub

Private Sub lstTables_Click()
  Dim i As Integer
  Dim fld As Field

  ListItemNames gdbCurrentDB.TableDefs(StripConnect(lstTables.Text)).Fields, lstFields, True
End Sub

Private Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
  lblLabels(2).Caption = LABEL3
  lblLabels(3).Caption = LABEL4
  
  RefreshTables frmReplace.lstTables
End Sub

Private Sub cmdOK_Click()
  Dim sReplaceTxt As String       'cadena de reemplazo
  Dim sWhereCondition As String   'condición Where

  On Error GoTo ReplaceErr

  MsgBar MSG1, True
  If gdbCurrentDB.TableDefs(StripConnect(lstTables.Text)).Fields(lstFields).Type = dbText Then
    sReplaceTxt = "'" & txtReplaceWith & "'"
  Else
    sReplaceTxt = txtReplaceWith
  End If
  If Len(txtCondition.Text) = 0 Then
    sWhereCondition = vbNullString
  Else
    sWhereCondition = " where " & txtCondition
  End If
  If gsDataType = gsSQLDB Then
    gdbCurrentDB.Execute "update " & StripConnect(lstTables.Text) & " set " & lstFields.Text & "=" & sReplaceTxt & sWhereCondition, dbSQLPassThrough
  Else
    gdbCurrentDB.Execute "update " & StripConnect(lstTables.Text) & " set [" & lstFields.Text & "]=" & sReplaceTxt & sWhereCondition
  End If
  If gdbCurrentDB.RecordsAffected > 0 Then
    If gbTransPending Then gbDBChanged = True
  End If

  Unload Me

  MsgBar vbNullString, False
  Exit Sub

ReplaceErr:
  ShowError
End Sub
