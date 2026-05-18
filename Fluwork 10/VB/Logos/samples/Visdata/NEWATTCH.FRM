VERSION 5.00
Begin VB.Form frmNewAttach 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva tabla adjunta"
   ClientHeight    =   3465
   ClientLeft      =   4830
   ClientTop       =   5325
   ClientWidth     =   4245
   HelpContextID   =   2016135
   Icon            =   "NEWATTCH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboConnect 
      Height          =   315
      ItemData        =   "NEWATTCH.frx":030A
      Left            =   120
      List            =   "NEWATTCH.frx":030C
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.CheckBox chkExclusive 
      Caption         =   "Exclusiva"
      Height          =   225
      Left            =   2160
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   2520
      Width           =   1785
   End
   Begin VB.CheckBox chkSavePassword 
      Caption         =   "Guardar contraseña"
      Height          =   225
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   2520
      Width           =   1740
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   2160
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   2880
      Width           =   1905
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "&Adjuntar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   10
      Top             =   2880
      Width           =   1905
   End
   Begin VB.ComboBox cboTableName 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtAttachName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de la base de datos: "
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   720
      Width           =   2385
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Adjuntar la tabla: "
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Cadena de conexión: "
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   4
      Top             =   1320
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del adjunto: "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Nueva tabla adjunta"
Const BUTTON1 = "&Adjuntar"
Const BUTTON2 = "&Cancelar"
Const Label1 = "&Nombre del adjunto:"
Const Label2 = "Nombre de la &base de datos:"
Const LABEL3 = "Ca&dena de conexión:"
Const LABEL4 = "&Tabla para adjuntar:"
Const CHECK1 = "AttachSave&PWD"
Const CHECK2 = "Attach&Exclusive"
'>>>>>>>>>>>>>>>>>>>>>>>>

Dim mbAdded As Boolean

Private Sub cmdCancel_Click()
  If mbAdded Then
    'han agregado algunas, de modo que tenemos que actualizar la lista de tablas
    RefreshTables Nothing
  End If

  Unload Me
End Sub

Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdAttach.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
  lblLabels(2).Caption = LABEL3
  lblLabels(3).Caption = LABEL4
  chkSavePassword.Caption = CHECK1
  chkExclusive.Caption = CHECK2
  
  cboConnect.AddItem "MDB de Access"
  cboConnect.AddItem "Dbase III"
  cboConnect.AddItem "Dbase IV"
  cboConnect.AddItem "FoxPro 2.0"
  cboConnect.AddItem "FoxPro 2.5"
  cboConnect.AddItem "FoxPro 2.6"
  cboConnect.AddItem "Paradox 3.X"
  cboConnect.AddItem "Paradox 4.X"
  cboConnect.AddItem "Excel 5.0"
  cboConnect.AddItem "Excel 4.0"
  cboConnect.AddItem "Excel 3.0"
  cboConnect.AddItem "Archivos de texto"
  cboConnect.AddItem "ODBC"
  
  cboConnect.ListIndex = 0
  
End Sub

Private Sub txtAttachName_Change()
  cmdAttach.Enabled = Len(txtAttachName.Text) > 0
  If cboTableName.ListCount > 0 Then cboTableName.Clear
End Sub

Private Sub cboConnect_Change()
  cmdAttach.Enabled = Len(txtAttachName.Text) > 0
  If cboTableName.ListCount > 0 Then cboTableName.Clear
End Sub

Private Sub cboTableName_DropDown()
  Dim db As Database
  Dim i As Integer
  Dim sDataType As String       'cadena de tipo de datos
  Dim tdf As TableDef

  On Error GoTo DDErr

  Screen.MousePointer = vbHourglass
  If cboTableName.ListCount = 0 Then
    'llena la lista
    sDataType = GetConnectStr()
    
    If cboConnect.ListIndex = 0 Then
      'caso especial para tablas mdb de Microsoft Access
      Set db = gwsMainWS.OpenDatabase(vbNullString, 0, 0, sDataType)
    Else
      Set db = gwsMainWS.OpenDatabase(vbNullString, 0, 0, sDataType & ";" & cboConnect.Text)
    End If
    For Each tdf In db.TableDefs
      If (tdf.Attributes And dbSystemObject) = 0 Then
        cboTableName.AddItem tdf.Name
      End If
    Next
  End If

  Screen.MousePointer = vbDefault
  Exit Sub

DDErr:
  ShowError
End Sub

Private Sub cmdAttach_Click()
  frmAttachments.AddAttachment
  mbAdded = True
End Sub

Public Function GetConnectStr() As String
  On Error GoTo GCErr
  
  Dim sTmp As String
  
  If cboConnect.ListIndex = 0 Then
    sTmp = ";"
  Else
    sTmp = cboConnect.Text & ";"
  End If
  
  If Len(txtDatabase.Text) > 0 Then
    sTmp = sTmp & "database=" & txtDatabase
  End If
  
  GetConnectStr = sTmp
  
  Exit Function
  
GCErr:
  ShowError
End Function
