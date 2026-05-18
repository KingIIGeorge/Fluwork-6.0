VERSION 5.00
Begin VB.Form frmTblStruct 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estructura de la tabla"
   ClientHeight    =   6135
   ClientLeft      =   1560
   ClientTop       =   945
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016147
   Icon            =   "TBLSTRU.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picFieldProps 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3120
      ScaleHeight     =   615
      ScaleWidth      =   4455
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4335
      Width           =   4455
      Begin VB.CheckBox chkUnique 
         Caption         =   "Único"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         MaskColor       =   &H00000000&
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   0
         Width           =   1230
      End
      Begin VB.CheckBox chkRequiredInd 
         Caption         =   "Necesario"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   360
         Width           =   1230
      End
      Begin VB.CheckBox chkIgnoreNull 
         Caption         =   "Ignorar nulos"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         MaskColor       =   &H00000000&
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   360
         Width           =   1590
      End
      Begin VB.CheckBox chkPrimary 
         Caption         =   "Primario"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   1230
      End
      Begin VB.CheckBox chkForeign 
         Caption         =   "Externo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         MaskColor       =   &H00000000&
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   1230
      End
   End
   Begin VB.PictureBox picFieldProps2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   4800
      ScaleHeight     =   1815
      ScaleWidth      =   3255
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3255
      Begin VB.CheckBox chkRequired 
         Caption         =   "Necesario"
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkAllowZeroLen 
         Caption         =   "Permitir longitud cero"
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   0
         Width           =   1995
      End
      Begin VB.TextBox txtOrdinalPos 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtValidationText 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   3075
      End
      Begin VB.TextBox txtValidationRule 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   3075
      End
      Begin VB.TextBox txtDefaultValue 
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   1440
         Width           =   3075
      End
   End
   Begin VB.PictureBox picFieldProps1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4800
      ScaleHeight     =   1095
      ScaleWidth      =   3255
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   840
      Width           =   3255
      Begin VB.TextBox txtCollatingOrder 
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoInc 
         Caption         =   "Incremento automático"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
         Width           =   2115
      End
      Begin VB.CheckBox chkVariable 
         Caption         =   "Longitud variable"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
         Width           =   1755
      End
      Begin VB.CheckBox chkFixedField 
         Caption         =   "Longitud fija"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   1400
      End
      Begin VB.TextBox txtFieldSize 
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboFieldType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "TBLSTRU.frx":000C
         Left            =   0
         List            =   "TBLSTRU.frx":000E
         Style           =   1  'Simple Combo
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtIndexName 
      Height          =   285
      Left            =   4680
      TabIndex        =   14
      Top             =   3975
      Width           =   3195
   End
   Begin VB.TextBox txtFieldName 
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Left            =   3960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5055
      Width           =   3915
   End
   Begin VB.ListBox lstIndexes 
      Height          =   870
      Left            =   120
      TabIndex        =   11
      Top             =   4215
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddTable 
      Caption         =   "&Generar la tabla"
      Enabled         =   0   'False
      Height          =   375
      HelpContextID   =   2016147
      Left            =   480
      MaskColor       =   &H00000000&
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H00000000&
      TabIndex        =   17
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir estructura"
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00000000&
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdRemoveIndex 
      Caption         =   "Eli&minar índice"
      Height          =   375
      Left            =   1560
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   5115
      Width           =   1440
   End
   Begin VB.CommandButton cmdAddIndex 
      Caption         =   "&Agregar índice"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   5115
      Width           =   1440
   End
   Begin VB.ListBox lstFields 
      Height          =   2625
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddField 
      Caption         =   "Agregar &campo"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   3360
      Width           =   1440
   End
   Begin VB.CommandButton cmdRemoveField 
      Caption         =   "E&liminar campo"
      Height          =   375
      Left            =   1545
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   3360
      Width           =   1440
   End
   Begin VB.TextBox txtTableName 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nombre: "
      Height          =   195
      Index           =   24
      Left            =   3240
      TabIndex        =   31
      Top             =   3975
      Width           =   870
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nombre: "
      Height          =   195
      Index           =   20
      Left            =   3120
      TabIndex        =   30
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Campos: "
      Height          =   195
      Index           =   23
      Left            =   3240
      TabIndex        =   29
      Top             =   5055
      Width           =   690
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Valor predeterminado:"
      Height          =   195
      Index           =   10
      Left            =   3120
      TabIndex        =   28
      Top             =   3435
      Width           =   1605
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Regla de validación:"
      Height          =   195
      Index           =   9
      Left            =   3120
      TabIndex        =   27
      Top             =   3075
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Texto de validación:"
      Height          =   195
      Index           =   8
      Left            =   3120
      TabIndex        =   26
      Top             =   2715
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Posición:"
      Height          =   195
      Index           =   7
      Left            =   3120
      TabIndex        =   25
      Top             =   2355
      Width           =   630
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Tamaño: "
      Height          =   195
      Index           =   5
      Left            =   3120
      TabIndex        =   24
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Tipo: "
      Height          =   195
      Index           =   4
      Left            =   3120
      TabIndex        =   23
      Top             =   840
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   7860
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Orden: "
      Height          =   195
      Index           =   22
      Left            =   3120
      TabIndex        =   22
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Lista de índices: "
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   3975
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Lista de campos: "
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de la tabla: "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmTblStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Estructura de la tabla"
Const BUTTON1 = "Ag&regar campo"
Const BUTTON2 = "E&liminar campo"
Const BUTTON3 = "&Agregar índice"
Const BUTTON4 = "&Eliminar índice"
Const BUTTON5 = "&Generar la tabla"
Const BUTTON6 = "&Cerrar"
Const BUTTON7 = "&Imprimir estructura"
Const Label1 = "&Nombre de la tabla:"
Const Label2 = "&Lista de campos:"
Const LABEL3 = "Lis&ta de índices:"
Const MSG1 = "Introduzca los nuevos parámetros del campo, presione 'Cerrar' al terminar"
Const MSG2 = "Introduzca los nuevos parámetros del índice, presione 'Cerrar' al terminar"
Const MSG3 = "Agregando la nueva tabla a la base de datos"
Const MSG4 = "¿Desea eliminar el índice?"
Const MSG5 = "Abriendo el diseño de formulario"
Const MSG6 = "Imprimiendo la estructura de la tabla"
Const MSG7 = "¿Desea eliminar el campo?"
Const MSG8 = "¿Desea cerrar sin guardar?"
'>>>>>>>>>>>>>>>>>>>>>>>>

Dim msCurrField As String
Dim mfldCurrFld As Field
Dim msCurrIndex As String
Dim mindCurrInd As Index
Dim mnFldCount As Integer
Dim mnIndCount As Integer
Dim mbTableNameChanged As Boolean

Sub cboFieldType_Change()
  If mfldCurrFld.Type < 9 Then
    cboFieldType.ListIndex = mfldCurrFld.Type - 1
  Else
    cboFieldType.ListIndex = mfldCurrFld.Type - 2
  End If
End Sub

Sub cboFieldType_Click()
  If cboFieldType.ListIndex = -1 Then Exit Sub
  If mfldCurrFld.Type < 9 Then
    cboFieldType.ListIndex = mfldCurrFld.Type - 1
  Else
    cboFieldType.ListIndex = mfldCurrFld.Type - 2
  End If
End Sub

Private Sub chkAllowZeroLen_Click()
  On Error GoTo AZErr
  
  If mfldCurrFld Is Nothing Then Exit Sub
  
  mfldCurrFld.AllowZeroLength = IIf(chkAllowZeroLen.Value = vbChecked, True, False)
  Exit Sub
  
AZErr:
  ShowError
End Sub

Private Sub chkRequired_Click()
  On Error GoTo RQErr
  If mfldCurrFld Is Nothing Then Exit Sub
  
  mfldCurrFld.Required = IIf(chkRequired.Value = vbChecked, True, False)
  Exit Sub
  
RQErr:
  ShowError
End Sub

Private Sub cmdAddField_Click()
  MsgBar MSG1, False
  frmAddField.Show vbModal
  MsgBar vbNullString, False
End Sub

Private Sub cmdAddIndex_Click()
  MsgBar MSG2, False
  frmAddIndex.Show vbModal
  MsgBar vbNullString, False
End Sub

Private Sub cmdAddTable_Click()
  On Error GoTo ATErr

  Dim i As Integer

  If DupeTableName(gtdfTableDef.Name) Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  MsgBar MSG3, True

  'agrega la Tabledef
  gdbCurrentDB.TableDefs.Append gtdfTableDef

  RefreshTables Nothing

  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  Unload Me
  Exit Sub

ATErr:
  ShowError
End Sub

Private Sub cmdClose_Click()
  If mbTableNameChanged Then
    RefreshTables Nothing
  End If
  If cmdAddTable.Visible And cmdAddTable.Enabled Then
    If MsgBox(MSG8, vbYesNo + vbQuestion, Me.Caption) = vbYes Then
      Unload Me
      MsgBar vbNullString, False
    End If
  Else
    Unload Me
    MsgBar vbNullString, False
  End If
End Sub

Sub lstFields_Click()
  On Error GoTo FErr

  If lstFields.ListIndex = -1 Then Exit Sub
  
  msCurrField = lstFields.Text
  Set mfldCurrFld = gtdfTableDef.Fields(msCurrField)
    
  'sólo activa estos campos si hay un campo actual en una base de datos Access
  txtFieldName.Enabled = (gsDataType = gsMSACCESS)
  txtValidationText.Enabled = (gsDataType = gsMSACCESS)
  txtValidationRule.Enabled = (gsDataType = gsMSACCESS)
  txtDefaultValue.Enabled = (gsDataType = gsMSACCESS)
  chkRequired.Enabled = (gsDataType = gsMSACCESS)
  chkAllowZeroLen.Enabled = (gsDataType = gsMSACCESS)
  txtOrdinalPos.Enabled = (gsDataType = gsMSACCESS)
  
  'desbloquea el campo de nombre
  txtFieldName.Locked = False
  txtFieldName.Text = mfldCurrFld.Name
  txtOrdinalPos.Text = mfldCurrFld.OrdinalPosition
  If mfldCurrFld.Type < 9 Then
    cboFieldType.ListIndex = mfldCurrFld.Type - 1
  Else
    cboFieldType.ListIndex = mfldCurrFld.Type - 2
  End If
  txtFieldSize.Text = mfldCurrFld.Size
  txtCollatingOrder.Text = mfldCurrFld.CollatingOrder
  chkFixedField.Value = IIf((mfldCurrFld.Attributes And dbFixedField) = dbFixedField, 1, 0)
  chkVariable.Value = IIf((mfldCurrFld.Attributes And dbVariableField) = dbVariableField, 1, 0)
  chkAutoInc.Value = IIf((mfldCurrFld.Attributes And dbAutoIncrField) = dbAutoIncrField, 1, 0)
  
  If gsDataType = gsMSACCESS Then
    txtValidationText.Text = mfldCurrFld.ValidationText
    txtValidationRule.Text = mfldCurrFld.ValidationRule
    txtDefaultValue.Text = mfldCurrFld.DefaultValue
    chkRequired.Value = IIf(mfldCurrFld.Required, 1, 0)
    chkAllowZeroLen.Value = IIf(mfldCurrFld.AllowZeroLength, 1, 0)
  End If
  
  Exit Sub
  
FErr:
  ShowError
End Sub

Sub lstIndexes_Click()
  On Error GoTo IErr

  If lstIndexes.ListIndex = -1 Then Exit Sub
  
  msCurrIndex = lstIndexes.Text
  Set mindCurrInd = gtdfTableDef.Indexes(msCurrIndex)
    
  txtIndexName.Text = mindCurrInd.Name
  txtFields.Text = mindCurrInd.Fields
  chkRequiredInd.Value = IIf(mindCurrInd.Required, 1, 0)
  chkUnique.Value = IIf(mindCurrInd.Unique, 1, 0)
  chkIgnoreNull.Value = IIf(mindCurrInd.IgnoreNulls, 1, 0)
  
  If gsDataType = gsMSACCESS Then
    chkPrimary.Value = IIf(mindCurrInd.Primary, 1, 0)
    chkForeign.Value = IIf(mindCurrInd.Foreign, 1, 0)
  End If
  
  Exit Sub
  
IErr:
  ShowError
End Sub

Private Sub txtCollatingOrder_LostFocus()
  If mfldCurrFld Is Nothing Then Exit Sub
  
  'lo restablece porque es de sólo lectura
  txtCollatingOrder.Text = mfldCurrFld.CollatingOrder
End Sub

Private Sub txtDefaultValue_LostFocus()
  On Error GoTo DVErr
  
  If mfldCurrFld Is Nothing Then Exit Sub
  
  If mfldCurrFld.DefaultValue <> txtDefaultValue.Text Then
    If Len(txtDefaultValue.Text) > 0 Then
      mfldCurrFld.DefaultValue = txtDefaultValue.Text
    End If
  End If
  Exit Sub
  
DVErr:
  ShowError
End Sub

Private Sub txtFieldName_LostFocus()
  On Error GoTo FNErr
  
  Dim i As Integer
  
  If mfldCurrFld Is Nothing Then Exit Sub
  
  'cambia el nombre si el usuario lo ha modificado
  If mfldCurrFld.Name <> txtFieldName.Text Then
    If Len(txtFieldName.Text) > 0 Then
      For i = 0 To lstFields.ListCount - 1
        If lstFields.List(i) = mfldCurrFld.Name Then
          lstFields.RemoveItem i
          lstFields.AddItem txtFieldName.Text, i
          Exit For
        End If
      Next
      mfldCurrFld.Name = txtFieldName.Text
    End If
  End If
  Exit Sub
  
FNErr:
  ShowError
End Sub

Sub txtFields_LostFocus()
  If mindCurrInd Is Nothing Then Exit Sub
  
  'lo restablece porque es de sólo lectura
  txtFields.Text = mindCurrInd.Fields
End Sub

Private Sub txtFieldSize_LostFocus()
  If mfldCurrFld Is Nothing Then Exit Sub
  
  'lo restablece porque es de sólo lectura
  txtFieldSize.Text = mfldCurrFld.Size
End Sub

Private Sub txtIndexName_LostFocus()
  On Error GoTo IDNErr
  
  Dim i As Integer
  
  If mindCurrInd Is Nothing Then Exit Sub
  
  'cambia el nombre si el usuario lo ha modificado
  If mindCurrInd.Name <> txtIndexName.Text Then
    If Len(txtIndexName.Text) > 0 And gsDataType = gsMSACCESS Then
      For i = 0 To lstIndexes.ListCount - 1
        If lstIndexes.List(i) = mindCurrInd.Name Then
          lstIndexes.RemoveItem i
          lstIndexes.AddItem txtIndexName.Text, i
          Exit For
        End If
      Next
      mindCurrInd.Name = txtIndexName.Text
    End If
  End If
  Exit Sub
  
IDNErr:
  ShowError
End Sub

Private Sub txtOrdinalPos_LostFocus()
  On Error GoTo OPErr
  
  If mfldCurrFld Is Nothing Then Exit Sub
  
  If mfldCurrFld.OrdinalPosition <> txtOrdinalPos.Text Then
    If Len(txtFieldName.Text) > 0 And gsDataType = gsMSACCESS Then
      mfldCurrFld.OrdinalPosition = txtOrdinalPos.Text
    End If
  End If
  Exit Sub
  
OPErr:
  ShowError
End Sub

Private Sub txtTableName_Change()
  If gbAddTableFlag Then
    If Len(txtTableName.Text) > 0 And lstFields.ListCount > 0 Then
      cmdAddTable.Enabled = True
    Else
      cmdAddTable.Enabled = False
    End If
    gtdfTableDef.Name = txtTableName.Text
  End If
End Sub

Private Sub txtTableName_LostFocus()
  On Error GoTo TBNErr
  
  Dim i As Integer
  
  'cambia el nombre si el usuario lo ha modificado
  If gtdfTableDef.Name <> txtTableName.Text Then
    If Len(txtTableName.Text) > 0 And gsDataType = gsMSACCESS Then
      'busca y cambia el nombre de la entrada en la lista del formulario de las tablas
      gtdfTableDef.Name = txtTableName.Text
      mbTableNameChanged = True
    End If
  End If
  Exit Sub
  
TBNErr:
  ShowError
End Sub

Private Sub txtTableName_KeyPress(KeyAscii As Integer)
  If txtTableName.TabStop = False Then
    KeyAscii = 0   'se deshace de la tecla
  End If
End Sub

Private Sub cmdRemoveIndex_Click()
  On Error GoTo DELErr

  If lstIndexes.ListIndex < 0 Then Exit Sub
  
  If MsgBox(MSG4, vbYesNo + vbQuestion) = vbYes Then
    If gbAddTableFlag = False Then
      gtdfTableDef.Indexes.Delete lstIndexes.Text
    End If
    'actualiza la lista de índices
    lstIndexes.RemoveItem lstIndexes.ListIndex
  End If
  
  'borra las propiedades
  txtIndexName.Text = vbNullString
  txtFields.Text = vbNullString
  chkRequiredInd.Value = vbUnchecked
  chkUnique.Value = vbUnchecked
  chkIgnoreNull.Value = vbUnchecked
  chkPrimary.Value = vbUnchecked
  chkForeign.Value = vbUnchecked
  
  Exit Sub

DELErr:
  ShowError
End Sub

Private Sub Form_Load()
  On Error GoTo LoadErr
  
  Dim fld As Field
  Dim idx As Index
  
  Me.Caption = FORMCAPTION
  cmdAddField.Caption = BUTTON1
  cmdRemoveField.Caption = BUTTON2
  cmdAddIndex.Caption = BUTTON3
  cmdRemoveIndex.Caption = BUTTON4
  cmdAddTable.Caption = BUTTON5
  cmdClose.Caption = BUTTON6
  cmdPrint.Caption = BUTTON7
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
  lblLabels(2).Caption = LABEL3
  
  Screen.MousePointer = vbHourglass
  MsgBar MSG5, True
  
  cboFieldType.AddItem "Boolean"
  cboFieldType.AddItem "Byte"
  cboFieldType.AddItem "Integer"
  cboFieldType.AddItem "Long"
  cboFieldType.AddItem "Currency"
  cboFieldType.AddItem "Single"
  cboFieldType.AddItem "Double"
  cboFieldType.AddItem "Date/Time"
  cboFieldType.AddItem "Text"
  cboFieldType.AddItem "Binary"
  cboFieldType.AddItem "Memo"
  
  If gbAddTableFlag Then
    Set gtdfTableDef = gdbCurrentDB.CreateTableDef()
    mnFldCount = 0
    mnIndCount = 0
    cmdAddTable.Visible = True
  Else
    cmdPrint.Visible = True
    Set gtdfTableDef = gdbCurrentDB.TableDefs(StripConnect(gnodDBNode2.Text))
    txtTableName.Text = gtdfTableDef.Name
    ListItemNames gtdfTableDef.Fields, lstFields, False
    mnFldCount = lstFields.ListCount
    lstFields.ListIndex = 0
    ListItemNames gtdfTableDef.Indexes, lstIndexes, False
    mnIndCount = lstIndexes.ListCount
    If mnIndCount > 0 Then lstIndexes.ListIndex = 0
  End If
  
  If gsDataType <> gsMSACCESS Then
    'no puede cambiar el nombre de las tablas en las bases de datos no mdb
    If gbAddTableFlag = False Then txtTableName.Locked = True
    'no puede eliminar campos en las tablas no mdb
    If gbAddTableFlag = False Then cmdRemoveField.Enabled = False
    'deshabilita las propiedades que no se pueden cambiar en las tablas no mdb
    txtFieldName.Locked = True
    chkRequired.Enabled = False
    chkAllowZeroLen.Enabled = False
    
    txtIndexName.Locked = True
    txtFields.Locked = True
  End If
  
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  Exit Sub

LoadErr:
  ShowError
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo PRTErr
  
  'esta subrutina simplemente imprime la definición
  'de la tabla actualmente seleccionada

  Dim i As Integer
  Dim sTmp As String

  MsgBar MSG6, True
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print "Base de datos: " & gsDBName
  Printer.Print
  Printer.Print
  Printer.Print "Definición de la tabla de " & txtTableName
  Printer.Print
  Printer.Print
  Printer.Print "Campos: (Nombre - Tipo - Tamaño)"
  Printer.Print String(60, "-")
  For i = 0 To lstFields.ListCount - 1
    lstFields.ListIndex = i
    sTmp = txtFieldName.Text & " - "
    sTmp = sTmp & cboFieldType.Text & " - "
    sTmp = sTmp & txtFieldSize.Text
    Printer.Print sTmp
  Next
  Printer.Print
  Printer.Print
  Printer.Print "Índices: (Nombre - Tipo - Tamaño)"
  Printer.Print String(60, "-")
  For i = 0 To lstIndexes.ListCount - 1
    sTmp = txtIndexName.Text & " - "
    sTmp = sTmp & txtFields.Text & " - "
    sTmp = sTmp & IIf(chkUnique = 1, "True", "False")
    Printer.Print sTmp
  Next
  Printer.NewPage
  Printer.EndDoc
  MsgBar vbNullString, False
  Exit Sub
  
PRTErr:
  ShowError
End Sub

Private Sub cmdRemoveField_Click()
  On Error GoTo RFErr

  If lstFields.ListIndex < 0 Then Exit Sub

  If MsgBox(MSG7, vbYesNo + vbQuestion) = vbYes Then
    'borra los valores de las propiedades del campo
    txtFieldName.Text = vbNullString
    txtOrdinalPos.Text = vbNullString
    cboFieldType.ListIndex = -1
    cboFieldType.Text = vbNullString
    txtFieldSize.Text = vbNullString
    txtCollatingOrder.Text = vbNullString
    chkFixedField.Value = vbUnchecked
    chkVariable.Value = vbUnchecked
    chkAutoInc.Value = vbUnchecked
    txtValidationText.Text = vbNullString
    txtValidationRule.Text = vbNullString
    txtDefaultValue.Text = vbNullString
    chkRequired.Value = vbUnchecked
    chkAllowZeroLen.Value = vbUnchecked
    'elimina de la estructura Tabledef
    gtdfTableDef.Fields.Delete lstFields.Text
    'elimina de mi lista
    lstFields.RemoveItem lstFields.ListIndex
  End If
  If lstFields.ListCount = 0 Then
    'no hay campos, luego desactiva el botón Generar
    cmdAddTable.Enabled = False
  End If
  
  Exit Sub

RFErr:
  ShowError
End Sub


Private Sub txtValidationRule_LostFocus()
  On Error GoTo VRErr
  
  If mfldCurrFld Is Nothing Then Exit Sub
  
  If mfldCurrFld.ValidationRule <> txtValidationRule.Text Then
    If Len(txtValidationRule.Text) > 0 And gsDataType = gsMSACCESS Then
      mfldCurrFld.ValidationRule = txtValidationRule.Text
    End If
  End If
  Exit Sub
  
VRErr:
  ShowError
End Sub

Private Sub txtValidationText_LostFocus()
  On Error GoTo VTErr
  
  If mfldCurrFld Is Nothing Then Exit Sub
  
  If mfldCurrFld.ValidationText <> txtValidationText.Text Then
    If Len(txtValidationText.Text) > 0 And gsDataType = gsMSACCESS Then
      mfldCurrFld.ValidationText = txtValidationText.Text
    End If
  End If
  Exit Sub
  
VTErr:
  ShowError
End Sub
