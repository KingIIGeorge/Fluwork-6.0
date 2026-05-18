VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmDataControl 
   ClientHeight    =   2160
   ClientLeft      =   1230
   ClientTop       =   3075
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016122
   Icon            =   "DATAFORM.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Tag             =   "Recordset"
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   0
      Width           =   5772
      Begin VB.CommandButton cmdCancelAdd 
         Caption         =   "Ca&ncelar"
         Height          =   330
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Renovar"
         Height          =   330
         Left            =   3840
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Ca&mpo"
         Height          =   330
         Left            =   2880
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   330
         Left            =   4800
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   1920
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         Height          =   330
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Actualizar"
         Height          =   330
         Left            =   960
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Top             =   0
         Width           =   960
      End
      Begin VB.Label lblFieldHeader 
         Caption         =   "Nombre del campo:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblFieldValue 
         Caption         =   " Valor:"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Data datDataCtl 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Tag             =   "OLE"
      Top             =   1815
      Width           =   5772
   End
   Begin VB.VScrollBar vsbScrollBar 
      Height          =   2440
      LargeChange     =   3000
      Left            =   7665
      SmallChange     =   300
      TabIndex        =   11
      Top             =   630
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picFields 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1056.479
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   7600.264
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   7605
      Begin VB.TextBox txtFieldData 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "datDataCtl"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1665
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox chkFieldData 
         DataSource      =   "datDataCtl"
         Height          =   282
         Index           =   0
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   735
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.PictureBox picFieldData 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "datDataCtl"
         Height          =   285
         Index           =   0
         Left            =   1680
         ScaleHeight     =   225
         ScaleWidth      =   3210
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.OLE oleFieldData 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "datDataCtl"
         DisplayType     =   1  'Icon
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblFieldName 
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   2415
      Top             =   1755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FilterIndex     =   1019
      FontSize        =   1.17491e-38
   End
End
Attribute VB_Name = "frmDataControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const BUTTON1 = "Ca&ncelar"
Const BUTTON2 = "A&ctualizar"
Const BUTTON3 = "&Eliminar"
Const BUTTON4 = "Ca&mpo"
Const BUTTON5 = "&Renovar"
Const BUTTON6 = "C&errar"
Const BUTTON7 = "&Agregar"
Const Label1 = "Nombre del campo:"
Const Label2 = "Valor:"
Const MSG1 = "Nuevo registro"
Const MSG2 = "Seleccione un archivo de imagen para cargar"
Const MSG3 = "El evento error de datos generó el error:"
Const MSG4 = "¿Desea guardar el nuevo registro?"
Const MSG5 = "¿Desea realizar los cambios?"
Const MSG6 = "¿Desea realizar los cambios antes de salir?"
Const MSG7 = "¿Desea eliminar el registro actual?"
Const MSG8 = "Escriba el valor de búsqueda:"
Const MSG9 = "Escriba la expresión de búsqueda:"
Const MSG10 = " Filas"
Const MSG11 = " Filas (Recordset de tipo Forward Only)"
Const MSG12 = "  [No actualizable]"
Const MSG13 = "Función no disponible para una consulta con parámetros"
'>>>>>>>>>>>>>>>>>>>>>>>>

'============================================================================
' Este es un formulario bastante genérico que puede ser utilizado en
' la mayor parte de los casos con una Table o una Querydef.
'============================================================================

Dim maFldArr() As Object
Public mrsFormRecordset As Recordset
Public mbIsParameterized As Boolean
Dim mvBookMark As Variant        'marcador del formulario
Dim mnNumFields As Integer       'número de campos
Dim mlNumRows As Long            'número de registros del Recordset
Dim mbJustUsedFind As Boolean    'indicador de la función de búsqueda
Dim mbResizing As Boolean        'indicador para evitar la recursividad de cambio de tamaño
Dim mbCancel As Boolean          'indicador para cancelar un Addnew
Dim mnFieldTop As Integer        'posición del primer campo

Const mnMSGBOX_TYPE = vbYesNo + vbQuestion
Const mnCTLARRAYHEIGHT = 340

Private Sub cmdAdd_Click()
  On Error GoTo AddErr

  datDataCtl.Recordset.AddNew
  datDataCtl.Caption = MSG1
  cmdCancelAdd.Visible = True
  cmdAdd.Visible = False
  If datDataCtl.Recordset.RecordCount <> 0 Then
    mvBookMark = datDataCtl.Recordset.Bookmark
    maFldArr(0).SetFocus
  End If

  Exit Sub
  
AddErr:
  ShowErrMsg

End Sub

Private Sub cmdCancelAdd_Click()
  On Error Resume Next

  mbCancel = True
  'vuelve al registro actual anterior
  If Len(mvBookMark) > 0 Then
    datDataCtl.Recordset.Bookmark = mvBookMark
  End If

End Sub

'------------------------------------------------------------
'esta subrutina carga la hoja de propiedades del control Data
'coméntela si va a hacer un uso independiente
'------------------------------------------------------------
Sub datDataCtl_MouseUp(BUTTON As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo DCPErr

  Dim i As Integer
  Dim recClone As Recordset
  Dim sTmpRS As String
  Dim sTmpDB As String
  Dim sTmpTag As String
  
  If BUTTON = 2 Then
    If mbIsParameterized Then
      MsgBox MSG13, vbInformation
      Exit Sub
    End If
  
    Screen.MousePointer = vbHourglass
    sTmpRS = datDataCtl.RecordSource
    sTmpDB = datDataCtl.DatabaseName
    sTmpTag = datDataCtl.Tag
    Set gDataCtlObj = datDataCtl
    frmDataCtlProp.Show vbModal
    If Not gDataCtlObj Is Nothing Then
      'verifica si el Recordset, base de datos o asociación ha cambiado,
      'y borra los campos asociados si alguno ha cambiado
      If (sTmpRS <> gDataCtlObj.RecordSource Or sTmpDB <> gDataCtlObj.DatabaseName) _
         Or gDataCtlObj.Tag <> sTmpTag Then
        'borra los nombres de los campos y quita la asociación de los controles
        For i = 0 To mnNumFields - 1
          lblFieldName(i).Caption = vbNullString
          maFldArr(i).DataField = vbNullString
          maFldArr(i).Visible = False
        Next
      End If
      datDataCtl.Refresh
      If (sTmpRS <> gDataCtlObj.RecordSource Or sTmpDB <> gDataCtlObj.DatabaseName) _
         Or gDataCtlObj.Tag <> sTmpTag Then
        Set recClone = datDataCtl.Recordset.Clone()
        If recClone.BOF = False And (datDataCtl.Options And dbForwardOnly) = 0 Then
          recClone.MoveLast
          mlNumRows = recClone.RecordCount
        Else
          mlNumRows = 0
        End If
        recClone.Close
        LoadFields
        SetRecNum
      Else
        'tiene que actualizar la copia local del Recordset para
        ' que lo utilicen otras operaciones, como el acceso a las propiedades
        Set mrsFormRecordset = datDataCtl.Recordset
        'si es un Recordset de recorrido unidireccional, tenemos que llamar a SetRecNum
        'para presentar el título correcto en el control Data
        If (datDataCtl.Options And dbForwardOnly) = dbForwardOnly Then
          SetRecNum
        End If
      End If
      gbSettingDataCtl = False
      If gDataCtlObj.Tag <> sTmpTag Then
        Form_Resize    'necesario para establecer la propiedad Left en los controles recién cargados
      End If
    End If
    
  End If
  Exit Sub
  
DCPErr:
  ShowError
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DBEngine.Idle dbFreeLocks
End Sub

Private Sub oleFieldData_MouseUp(Index As Integer, BUTTON As Integer, Shift As Integer, x As Single, y As Single)
  If BUTTON <> 2 Then Exit Sub
  'presenta el cuadro de diálogo que permite
  'la entrada en el campo OLE
  oleFieldData(Index).InsertObjDlg
End Sub

Private Sub txtFieldData_KeyPress(Index As Integer, KeyAscii As Integer)
  'pasa al siguiente campo al recibir la tecla Intro
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
  End If
End Sub

Private Sub picFieldData_Click(Index As Integer)
  'alterna el tamaño de un control imagen
  'para que pueda ser visto o comprimido
  If picFieldData(Index).Height <= 280 Then
    picFieldData(Index).AutoSize = True
  Else
    picFieldData(Index).AutoSize = False
    picFieldData(Index).Height = 280
  End If
End Sub

Private Sub picFieldData_DblClick(Index As Integer)
  On Error GoTo PicErr

  With dlgCMD1
    .Filter = "Mapas de bits (*.bmp)|*.bmp|Iconos (*.ico)|*.ico|Metarchivos (*.wmf)|*.wmf|Todos los archivos (*.*)|*.*"
    .DialogTitle = MSG2
    .FilterIndex = 1
    .ShowOpen
  
    If Len(.FileName) > 0 Then
      picFieldData(Index).Picture = LoadPicture(.FileName)
    End If
  End With

  Exit Sub
  
PicErr:
  ShowErrMsg
  Exit Sub

End Sub

Private Sub cmdClose_Click()
  On Error Resume Next

  Unload Me
End Sub

Private Sub vsbScrollBar_Change()
  Dim nCurrVal As Integer

  nCurrVal = vsbScrollBar
  If (nCurrVal - mnFieldTop) Mod mnCTLARRAYHEIGHT = 0 Then
    picFields.Top = nCurrVal
  Else
    picFields.Top = ((nCurrVal - mnFieldTop) \ mnCTLARRAYHEIGHT) * mnCTLARRAYHEIGHT + mnFieldTop
  End If

End Sub

Private Sub datDataCtl_Error(DataErr As Integer, Response As Integer)
  If DataErr = 481 Then  'ignora el error de imagen incorrecta
    Response = vbDataErrContinue
  Else
    MsgBox MSG3 & Error(DataErr)
  End If
End Sub

Private Sub datDataCtl_RePosition()
  On Error GoTo RepErr

  Dim sBookMark As String
  Dim recClone As Recordset
  
  'si vamos a restablecer el control Data, tenemos que saltar esta subrutina
  If gbSettingDataCtl Then Exit Sub
  
  'si el Recordset está vacío y aún no estamos en el modo AddNew
  'tenemos que llamar a la subrutina cmdAdd_Click para ejecutar un AddNew
  If (datDataCtl.Recordset.RecordCount = 0) And _
     (datDataCtl.EditMode <> dbEditAdd) And _
      datDataCtl.Recordset.Updatable Then
    Call cmdAdd_Click
    Exit Sub
  End If

  SetRecNum
  
  Exit Sub
  
RepErr:
  ShowErrMsg
  Exit Sub

End Sub

Private Sub datDataCtl_Validate(Action As Integer, Save As Integer)
  On Error GoTo ValErr

  If mbCancel Then
    Save = False
    mbCancel = False
    Exit Sub
  End If

  'antes comprueba si hay un movimiento desde un registro nuevo o en edición
  If Action < 5 Then
    If Save Then       'los datos han cambiado
      If datDataCtl.EditMode = dbEditAdd Then
        If MsgBox(MSG4, mnMSGBOX_TYPE) = vbYes Then
          mlNumRows = mlNumRows + 1
        Else
          Save = False
        End If
      Else
        If MsgBox(MSG5, mnMSGBOX_TYPE) <> vbYes Then
          Save = False        'pierde los cambios
        End If
      End If
    End If
  End If

  Select Case Action
    Case vbDataActionMoveFirst
      'no hace nada
    Case vbDataActionMovePrevious
      'no hace nada
    Case vbDataActionMoveNext
      'no hace nada
    Case vbDataActionMoveLast
      ' no hace nada
    Case vbDataActionAddNew
      ' no hace nada
    Case vbDataActionUpdate
       'pasa al código del evento cmdUpdate_click
    Case vbDataActionDelete
      ' no hace nada
    Case vbDataActionFind
      'establece el indicador para su uso en el evento de reposición
      mbJustUsedFind = True
    Case vbDataActionBookmark
      'no hace nada
    Case vbDataActionClose, vbDataActionUnload
      If Save Then
        If MsgBox(MSG6, mnMSGBOX_TYPE) <> vbYes Then
          Save = False
        End If
      End If

  End Select

  Exit Sub

ValErr:
  ShowErrMsg
  Exit Sub
  
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DELErr

  If MsgBox(MSG7, mnMSGBOX_TYPE) = vbYes Then
    datDataCtl.Recordset.Delete
    mlNumRows = mlNumRows - 1
    datDataCtl.Recordset.MoveNext
    'sale de la condición EOF si se diera
    If datDataCtl.Recordset.RecordCount > 0 Then datDataCtl.Recordset.MoveLast
    maFldArr(0).SetFocus
  End If

  Exit Sub

DELErr:
  ShowErrMsg
  Exit Sub

End Sub

Private Sub cmdFind_Click()
  On Error GoTo FindErr

  Dim sBookMark As String
  Dim sFindStr As String

  If datDataCtl.Recordset.Type = dbOpenTable Then
    sFindStr = InputBox(MSG8)
  Else
    sFindStr = InputBox(MSG9)
  End If
  If Len(sFindStr) = 0 Then Exit Sub

  If datDataCtl.Recordset.RecordCount > 0 Then
    sBookMark = datDataCtl.Recordset.Bookmark
  End If

  If datDataCtl.Recordset.Type = dbOpenTable Then
    datDataCtl.Recordset.Seek "=", sFindStr
  Else
    datDataCtl.Recordset.FindFirst sFindStr
  End If

  'vuelve al registro antiguo si no se ha encontrado nada
  If datDataCtl.Recordset.NoMatch And Len(sBookMark) > 0 Then
    datDataCtl.Recordset.Bookmark = sBookMark
  End If

  maFldArr(0).SetFocus
  Exit Sub

FindErr:
  ShowErrMsg
  maFldArr(0).SetFocus
  Exit Sub

End Sub

Private Sub Form_Load()
  On Error GoTo LoadErr

  cmdCancelAdd.Caption = BUTTON1
  cmdUpdate.Caption = BUTTON2
  cmdDelete.Caption = BUTTON3
  cmdFind.Caption = BUTTON4
  cmdRefresh.Caption = BUTTON5
  cmdClose.Caption = BUTTON6
  cmdAdd.Caption = BUTTON7
  lblFieldHeader.Caption = Label1
  lblFieldValue.Caption = Label2
  
  'mrsFormRecordset es una variable pública a nivel de módulo
  'que tiene que ser establecida antes de mostrar el formulario
  
  
  With mrsFormRecordset
    If .Type = dbOpenTable Then
      'hay que establecer el índice
      If gdbCurrentDB.TableDefs(.Name).Indexes.Count > 0 Then
        .Index = gdbCurrentDB.TableDefs(.Name).Indexes(0).Name
      End If
    End If
    If .RecordCount > 0 Then
      'pasa al siguiente y después al anterior. Así obtiene la cuenta de registros
      .MoveNext
      .MovePrevious
    End If
  End With
  Set datDataCtl.Recordset = mrsFormRecordset
  
  Me.Width = 5868
  LoadFields
  Me.Show
  maFldArr(0).SetFocus
  Exit Sub

LoadErr:
  ShowErrMsg
  Unload Me
End Sub

Private Sub Form_Resize()
  On Error Resume Next

  If gbSettingDataCtl Then Exit Sub
  If mbResizing Then Exit Sub

  Dim nHeight As Integer
  Dim i As Integer
  Dim nTotalWidth As Integer

  mbResizing = True
  If Me.WindowState <> 1 And lblFieldName(0).Visible Then 'no minimizado
    'asegura que el formulario está alineado con un campo
    nHeight = Me.Height
    If (nHeight - 1320) Mod mnCTLARRAYHEIGHT <> 0 Then
      Me.Height = ((nHeight - 1280) \ mnCTLARRAYHEIGHT) * mnCTLARRAYHEIGHT + 1280
    End If
    'cambia el tamaño de la barra de estado
    datDataCtl.Top = Me.Height - 650
    'cambia el tamaño de la barra de desplazamiento
    vsbScrollBar.Height = Me.Height - 1300
    vsbScrollBar.Left = Me.Width - 360
    If mrsFormRecordset.Fields.Count > 10 Then
      picFields.Width = Me.Width - 260
      nTotalWidth = vsbScrollBar.Left - 20
    Else
      picFields.Width = Me.Width - 20
      nTotalWidth = Me.Width - 50
    End If
    picButtons.Width = Me.Width - 20
    'amplía los campos si es posible
    For i = 0 To mrsFormRecordset.Fields.Count - 1
      lblFieldName(i).Width = 0.3 * nTotalWidth - 100
      maFldArr(i).Left = lblFieldName(i).Width + 200
      If datDataCtl.Recordset.Fields(i).Type > 9 Then
        maFldArr(i).Width = 0.7 * nTotalWidth - 270
      End If
    Next
    lblFieldValue.Left = maFldArr(0).Left

  End If

  mbResizing = False

End Sub

Private Function GetFieldWidth(rnType As Integer)
  'determina la anchura del control formulario
  'según el tipo de campo
  Select Case rnType
    Case dbBoolean
      GetFieldWidth = 850
    Case dbByte
      GetFieldWidth = 650
    Case dbInteger
      GetFieldWidth = 900
    Case dbLong
      GetFieldWidth = 1100
    Case dbCurrency
      GetFieldWidth = 1800
    Case dbSingle
      GetFieldWidth = 1800
    Case dbDouble
      GetFieldWidth = 2200
    Case dbDate
      GetFieldWidth = 2000
    Case dbText
      GetFieldWidth = 3250
    Case dbMemo
      GetFieldWidth = 3250
    Case Else
      GetFieldWidth = 3250
  End Select

End Function

Private Sub LoadFields()
   Dim recTmp As Recordset
   Dim nFldType As Integer
   Dim i As Integer

   On Error GoTo LoadFieldsErr

   Set mrsFormRecordset = datDataCtl.Recordset
   Set recTmp = mrsFormRecordset

   'carga los controles en el formulario del Recordset
   mnNumFields = recTmp.Fields.Count
   ReDim maFldArr(mnNumFields) As Object
   lblFieldName(0).Visible = True
   nFldType = recTmp.Fields(0).Type
   If nFldType = dbBoolean Then
     Set maFldArr(0) = chkFieldData(0)
   ElseIf nFldType = dbLongBinary Then
     If datDataCtl.Tag = "OLE" Then
       Set maFldArr(0) = oleFieldData(0)
     Else
       Set maFldArr(0) = picFieldData(0)
     End If
   Else
     Set maFldArr(0) = txtFieldData(0)
   End If
   maFldArr(0).Visible = True
   maFldArr(0).Top = 0
   maFldArr(0).Width = GetFieldWidth(nFldType)
   If nFldType = dbText Then maFldArr(0).MaxLength = recTmp.Fields(0).Size

   maFldArr(0).TabIndex = 0
   On Error Resume Next
   For i = 1 To recTmp.Fields.Count - 1
     picFields.Height = picFields.Height + mnCTLARRAYHEIGHT
     Load lblFieldName(i)
     lblFieldName(i).Top = lblFieldName(i - 1).Top + mnCTLARRAYHEIGHT
     lblFieldName(i).Visible = True
     nFldType = recTmp.Fields(i).Type
     If nFldType = dbBoolean Then
       Load chkFieldData(i)
       Set maFldArr(i) = chkFieldData(i)
     ElseIf nFldType = dbLongBinary Then
       If datDataCtl.Tag = "OLE" Then
         Load oleFieldData(i)
         Set maFldArr(i) = oleFieldData(i)
       Else
         Load picFieldData(i)
         Set maFldArr(i) = picFieldData(i)
       End If
     Else
       Load txtFieldData(i)
       Set maFldArr(i) = txtFieldData(i)
     End If
     maFldArr(i).Top = maFldArr(i - 1).Top + mnCTLARRAYHEIGHT
     maFldArr(i).Visible = True
     maFldArr(i).Width = GetFieldWidth(nFldType)
     maFldArr(i).TabIndex = i
     If nFldType = dbText Then maFldArr(i).MaxLength = recTmp.Fields(i).Size
   Next

   On Error GoTo LoadFieldsErr

   'cambia el tamaño de la ventana principal
   picFields.Top = picButtons.Top + picButtons.Height
   mnFieldTop = picFields.Top
   vsbScrollBar.Value = mnFieldTop
   If i <= 10 Then
     Height = i * mnCTLARRAYHEIGHT + 1500
     vsbScrollBar.Visible = False
   Else
     Height = 4500
     Width = Width + 260
     vsbScrollBar.Visible = True
     vsbScrollBar.Min = mnFieldTop
     vsbScrollBar.Max = mnFieldTop - (i * mnCTLARRAYHEIGHT) + 3000
   End If

   'presenta los nombres de los campos
   For i = 0 To recTmp.Fields.Count - 1
     lblFieldName(i).Caption = recTmp.Fields(i).Name & ":"
   Next
   
   'asocia los controles
   On Error Resume Next   'asocia aunque la tabla esté vacía
   For i = 0 To recTmp.Fields.Count - 1
     maFldArr(i).DataField = recTmp.Fields(i).Name
   Next

   Exit Sub

LoadFieldsErr:
   ShowErrMsg
   Exit Sub

End Sub

Private Sub cmdRefresh_Click()
  On Error GoTo RefErr

  If mbIsParameterized Then
    MsgBox MSG13, vbInformation
    Exit Sub
  End If
  
  datDataCtl.Refresh
  Exit Sub

RefErr:
  ShowErrMsg
End Sub

Private Sub SetRecNum()
  On Error GoTo SRErr
  
  Dim sCurrStat As String
  Dim lCurrRec As Long
  Dim bNoInd As Integer
  
  'obtiene la cuenta actual de registros
  mlNumRows = datDataCtl.Recordset.RecordCount
  
  If datDataCtl.EditMode <> dbEditAdd Then
    If datDataCtl.Recordset.BOF Then
      sCurrStat = "(BOF)/" & mlNumRows
    ElseIf datDataCtl.Recordset.EOF Then
      sCurrStat = "(EOF)/" & mlNumRows
    Else
      'comprueba si se utiliza una tabla con o sin índices
      If datDataCtl.Recordset.Type = dbOpenTable Then
        If datDataCtl.Database(datDataCtl.RecordSource).Indexes.Count = 0 Then
          bNoInd = True
        End If
      End If
      'si no hay índices en una tabla o si el Recordset es ForwardOnly
      'la propiedad PercentPosition no está disponible
      If bNoInd Then
        sCurrStat = mlNumRows & MSG10
      ElseIf (datDataCtl.Options And dbForwardOnly) = dbForwardOnly Then
        sCurrStat = mlNumRows & MSG11
      Else
        lCurrRec = (mlNumRows * (datDataCtl.Recordset.PercentPosition * 0.01)) + 1
        sCurrStat = lCurrRec & "/" & mlNumRows
      End If
    End If
    If datDataCtl.Recordset.Updatable = False Then
      sCurrStat = sCurrStat & MSG12
      cmdAdd.Enabled = False
      cmdCancelAdd.Enabled = False
      cmdUpdate.Enabled = False
      cmdDelete.Enabled = False
    Else
      cmdAdd.Enabled = True
      cmdCancelAdd.Enabled = True
      cmdUpdate.Enabled = True
      cmdDelete.Enabled = True
    End If
    datDataCtl.Caption = sCurrStat
  End If

  'restablece los botones si es necesario
  If datDataCtl.EditMode <> dbEditAdd Then
    cmdCancelAdd.Visible = False
    cmdAdd.Visible = True
  End If

  Exit Sub
SRErr:
  If Err <> 3021 Then
    ShowErrMsg
  End If
  Exit Sub

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdErr

  Dim bAddFlag As Integer

  Dim nDelay As Long
  Dim nRetryCnt As Integer

  bAddFlag = datDataCtl.EditMode
  
  If datDataCtl.EditMode = dbEditAdd Then
    If MsgBox(MSG4, mnMSGBOX_TYPE) = vbYes Then
      Screen.MousePointer = vbHourglass
RetryUpd1:
      datDataCtl.UpdateRecord
      mlNumRows = mlNumRows + 1
    End If
  Else
    If MsgBox(MSG5, mnMSGBOX_TYPE) = vbYes Then
      Screen.MousePointer = vbHourglass
RetryUpd2:
      datDataCtl.UpdateRecord
    End If
  End If

  If bAddFlag = dbEditAdd Then
    mrsFormRecordset.MoveLast
  End If

  DBEngine.Idle dbFreeLocks
  Screen.MousePointer = vbDefault
  Exit Sub

UpdErr:
  If Err = 3260 And nRetryCnt < gnMURetryCnt Then
    nRetryCnt = nRetryCnt + 1
    datDataCtl.Recordset.Bookmark = datDataCtl.Recordset.Bookmark   'Cancela la actualización
    DBEngine.Idle dbFreeLocks
    nDelay = Timer
    'Espera gnMUDelay segundos
    While Timer - nDelay < gnMUDelay
      'no hace nada
    Wend
    If datDataCtl.EditMode = dbEditAdd Then
      Resume RetryUpd1
    Else
      Resume RetryUpd2
    End If
  Else
    Screen.MousePointer = vbDefault
    ShowErrMsg
    Exit Sub
  End If

End Sub

Private Sub ShowErrMsg()
  MsgBox "Error:" & Err & " " & Error
End Sub
