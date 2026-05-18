VERSION 5.00
Begin VB.Form frmAddField 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar campo"
   ClientHeight    =   3735
   ClientLeft      =   2490
   ClientTop       =   2865
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016117
   Icon            =   "ADDFIELD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3297.239
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   6054.501
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optVariable 
      Caption         =   "Campo variable"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   2160
      Value           =   -1  'True
      Width           =   2379
   End
   Begin VB.OptionButton optFixedField 
      Caption         =   "Campo fijo"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   1920
      Width           =   2379
   End
   Begin VB.CheckBox chkAutoInc 
      Caption         =   "Incremento automático"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2379
   End
   Begin VB.CheckBox chkAllowZeroLen 
      Caption         =   "Permitir longitud cero"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   3000
      Width           =   2379
   End
   Begin VB.TextBox txtFieldName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.ComboBox cboFieldType 
      Height          =   315
      ItemData        =   "ADDFIELD.frx":030A
      Left            =   120
      List            =   "ADDFIELD.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtFieldSize 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtOrdinalPos 
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtValidationText 
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtValidationRule 
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtDefaultValue 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CheckBox chkRequired 
      Caption         =   "Necesario"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   3360
      Width           =   2379
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Nombre: "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Tipo: "
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Tamaño: "
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Posición:"
      Height          =   195
      Index           =   4
      Left            =   3120
      TabIndex        =   17
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Texto de validación:"
      Height          =   195
      Index           =   5
      Left            =   3120
      TabIndex        =   16
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Regla de validación:"
      Height          =   195
      Index           =   6
      Left            =   3120
      TabIndex        =   15
      Top             =   1380
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Valor predeterminado:"
      Height          =   195
      Index           =   7
      Left            =   3120
      TabIndex        =   14
      Top             =   2040
      Width           =   1605
   End
End
Attribute VB_Name = "frmAddField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Agregar campo"
Const BUTTON1 = "&Aceptar"
Const BUTTON2 = "&Cerrar"
Const MSG1 = " Ya existe"
'>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub cmdClose_Click()
  Unload Me
End Sub

Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdClose.Caption = BUTTON2

  cboFieldType.AddItem "Boolean"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbBoolean
  cboFieldType.AddItem "Byte"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbByte
  cboFieldType.AddItem "Integer"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbInteger
  cboFieldType.AddItem "Long"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbLong
  cboFieldType.AddItem "Currency"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbCurrency
  cboFieldType.AddItem "Single"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbSingle
  cboFieldType.AddItem "Double"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbDouble
  cboFieldType.AddItem "Date/Time"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbDate
  cboFieldType.AddItem "Text"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbText
  cboFieldType.AddItem "Binary"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbLongBinary
  cboFieldType.AddItem "Memo"
  cboFieldType.ItemData(cboFieldType.NewIndex) = dbMemo
  
  SetDefaults
  'es necesario desactivar los controles no aplicados a
  'tablas de Microsoft Access
  If gsDataType <> gsMSACCESS Then
    optFixedField.Enabled = False
    chkAutoInc.Enabled = False
    optVariable.Enabled = False
    txtValidationText.Enabled = False
    txtValidationRule.Enabled = False
    txtDefaultValue.Enabled = False
    chkRequired.Enabled = False
    chkAllowZeroLen.Enabled = False
  End If
  
End Sub

Private Sub txtFieldName_Change()
  'sólo activa el botón Aceptar si el
  'campo de nombre no está vacío
  cmdOK.Enabled = (Len(txtFieldName.Text) > 0)
End Sub

Private Sub cboFieldType_Click()
  Dim nFldType As Integer
  
  'llama a la función para establecer el tamaño y el tipo del campo
  txtFieldSize.Text = SetFldProperties(cboFieldType.ItemData(cboFieldType.ListIndex))
  txtFieldSize.Enabled = False

  nFldType = cboFieldType.ItemData(cboFieldType.ListIndex)
  
  'activa los controles adecuados para cada tipo de campo
  If gsDataType <> gsMSACCESS Then
    If nFldType = dbText Then
      'permite la introducción de la longitud del campo
      txtFieldSize.Enabled = True
      'tamaño predeterminado de campo de la IU de Access
      txtFieldSize.Text = "50"
    End If
    'el código siguiente sólo se ejecuta para bases de datos MDB
    Exit Sub
  End If
  
  If nFldType = dbText Then
    'permite la introducción de la longitud del campo
    txtFieldSize.Enabled = True
    'tamaño predeterminado de campo de la IU de Access
    txtFieldSize.Text = "50"
    'disponible para memo y texto
    chkAllowZeroLen.Enabled = True
    'sólo disponible para texto
    optVariable.Enabled = True
    optFixedField.Enabled = True
    'desactiva estos controles
    chkAutoInc.Enabled = False
    chkAutoInc.Value = vbUnchecked
  ElseIf nFldType = dbMemo Then
    'disponible para memo y texto
    chkAllowZeroLen.Enabled = True
    'desactiva estos controles
    optVariable.Enabled = False
    optFixedField.Enabled = False
    chkAutoInc.Enabled = False
    'establece el valor a 0
    optVariable.Value = False
    optFixedField.Value = False
    chkAutoInc.Value = vbUnchecked
  ElseIf nFldType = dbLong Then
    'activa éste para los campos de tipo contador
    chkAutoInc.Enabled = True
    'desactiva estos controles
    chkAllowZeroLen.Enabled = False
    optVariable.Enabled = False
    optFixedField.Enabled = False
    'establece el valor a 0
    chkAllowZeroLen.Value = vbUnchecked
    optVariable.Value = False
    optFixedField.Value = False
  Else
    'desactiva estos para los demás tipos
    chkAllowZeroLen.Enabled = False
    optVariable.Enabled = False
    optFixedField.Enabled = False
    chkAutoInc.Enabled = False
    'establece el valor a 0
    chkAllowZeroLen.Value = vbUnchecked
    optVariable.Value = False
    optFixedField.Value = False
    chkAutoInc.Value = vbUnchecked
  End If

End Sub

Private Sub cmdOK_Click()
  On Error GoTo OkayErr

  Dim fld As Field     'estructura de campo local
  Dim i As Integer

  'obtiene un objeto campo nuevo
  Set fld = gtdfTableDef.CreateField()
  'asigna valores a la estructura del campo
  With fld
    .Name = txtFieldName.Text
    .Type = cboFieldType.ItemData(cboFieldType.ListIndex)
    .Size = txtFieldSize.Text
    If Len(txtOrdinalPos.Text) > 0 Then .OrdinalPosition = txtOrdinalPos.Text
    
    If gsDataType = gsMSACCESS Then
      .Required = IIf(chkRequired.Value = vbChecked, -1, 0)
      If .Type = dbText Then
        'ésto sólo se aplica al texto
        .AllowZeroLength = IIf(chkAllowZeroLen.Value = vbChecked, -1, 0)
      End If
      If optFixedField.Value Then
        .Attributes = .Attributes Or dbFixedField
      End If
      If .Type = dbLong Then
        'sólo se aplica al tipo entero largo
        If chkAutoInc.Value = vbChecked Then
          .Attributes = .Attributes Or dbAutoIncrField
        End If
      End If
      If optVariable.Value Then
        .Attributes = .Attributes Or dbVariableField
      End If
      .ValidationText = txtValidationText.Text
      .ValidationRule = txtValidationRule.Text
      .DefaultValue = txtDefaultValue.Text
    End If
  End With
  
  'comprueba si hay duplicados
  If ObjectExists(gtdfTableDef.Fields, fld.Name) Then
    MsgBox "'" & fld.Name & "'" & MSG1
    txtFieldName.SelStart = 0
    txtFieldName.SelLength = Len(txtFieldName.Text)
    txtFieldName.SetFocus
    Exit Sub
  End If
  
  'intenta agregar el campo
  gtdfTableDef.Fields.Append fld
  
  'tiene que haber sido correcto, de modo que ...
  'agrega el elemento a la lista
  frmTblStruct.lstFields.AddItem txtFieldName
  'activa el nuevo elemento
  frmTblStruct.lstFields.ListIndex = frmTblStruct.lstFields.NewIndex

  'activa el botón de agregar tabla si es necesario
  If frmTblStruct.cmdAddTable.Visible Then
    frmTblStruct.cmdAddTable.Enabled = True
  End If
  
  'borra el nombre y permite la introducción de otro
  SetDefaults
  txtFieldName.SetFocus
  Exit Sub

OkayErr:
  ShowError
End Sub

Private Sub SetDefaults()
  txtFieldName.Text = vbNullString
  If gsDataType = gsMSACCESS Then
    optFixedField.Value = False
    chkAutoInc.Value = vbUnchecked
    optVariable.Value = True
    chkRequired.Value = vbUnchecked
    chkAllowZeroLen.Value = vbChecked
  Else
    optFixedField.Value = False
    chkAutoInc.Value = 2
    optVariable.Value = False
    chkRequired.Value = 2
    chkAllowZeroLen.Value = 2
  End If
  cboFieldType.ListIndex = 8             'de forma predeterminada, texto
  txtFieldSize.Text = 50                 'de forma predeterminada, 50
  txtValidationText.Text = vbNullString
  txtValidationRule.Text = vbNullString
  txtDefaultValue.Text = vbNullString
End Sub
