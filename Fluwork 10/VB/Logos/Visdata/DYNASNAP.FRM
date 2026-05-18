VERSION 5.00
Begin VB.Form frmDynaSnap 
   ClientHeight    =   3750
   ClientLeft      =   2730
   ClientTop       =   2610
   ClientWidth     =   5490
   HelpContextID   =   2016125
   Icon            =   "DYNASNAP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3733.906
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5503.698
   ShowInTaskbar   =   0   'False
   Tag             =   "Recordset"
   Begin VB.PictureBox picViewButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   5487.272
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   5484
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Mover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2730
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Ordenar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   372
         Width           =   1365
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filtrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1365
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Ce&rrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4095
         MaskColor       =   &H00000000&
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   15
         Width           =   1365
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Elimi&nar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2730
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   15
         Width           =   1365
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1365
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   15
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Top             =   20
         Width           =   1365
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Campo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4095
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   375
         Width           =   1365
      End
   End
   Begin VB.PictureBox picChangeButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   919.528
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   5719.056
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   960
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   48
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2640
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Top             =   48
         Width           =   1212
      End
   End
   Begin VB.PictureBox picFldHdr 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   14948.92
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   14946
      Begin VB.Label lblFieldValue 
         Caption         =   " Valor (F4=Zoom)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblFieldHdr 
         Caption         =   "Nombre del campo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.PictureBox picMoveButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   0
      ScaleHeight     =   298.153
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   5493.878
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3465
      Width           =   5484
      Begin VB.HScrollBar hsclCurrRow 
         Height          =   255
         Left            =   0
         Max             =   100
         TabIndex        =   9
         Top             =   29
         Width           =   2895
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   38
         Width           =   1695
      End
   End
   Begin VB.VScrollBar vsbScrollBar 
      Height          =   2250
      LargeChange     =   3000
      Left            =   5040
      SmallChange     =   300
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picFields 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   372
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   4812
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4815
      Begin VB.TextBox txtFieldData 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   3252
      End
      Begin VB.Label lblFieldName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   60
         Visible         =   0   'False
         Width           =   1572
      End
   End
End
Attribute VB_Name = "frmDynaSnap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const BUTTON1 = "&Agregar"
Const BUTTON2 = "&Editar"
Const BUTTON3 = "E&liminar"
Const BUTTON4 = "Ce&rrar"
Const BUTTON5 = "&Ordenar"
Const BUTTON6 = "&Filtrar"
Const BUTTON7 = "Mo&ver"
Const BUTTON8 = "Ca&mpo"
Const BUTTON9 = "&Cancelar"
Const BUTTON10 = "Actuali&zar"
Const Label1 = "Nombre del campo:"
Const Label2 = " Valor (F4=Zoom)"
Const MSG1 = "Agregar registro"
Const MSG2 = "Escriba el número de filas a mover:"
Const MSG3 = "(Use valores negativos para ir hacia atrás)"
Const MSG4 = "Se sobrepasó la longitud del campo, se truncaron los datos"
Const MSG5 = "¿Desea eliminar el registro actual?"
Const MSG6 = "Editar registro"
Const MSG7 = "Escriba la expresión de filtro:"
Const MSG8 = "Estableciendo un nuevo filtro"
Const MSG9 = "Escriba los parámetros de búsqueda"
Const MSG10 = "Buscando un nuevo registro"
Const MSG11 = "No se encontró el registro"
Const MSG12 = "Cambiando las dimensiones del formulario"
Const MSG13 = "Introduzca la columna de orden:"
Const MSG14 = "Estableciendo la nueva columna de orden"
'>>>>>>>>>>>>>>>>>>>>>>>>


'variables del formulario
Public mrsFormRecordset As Recordset
Dim msTableName As String      'nombre de la tabla
Dim mvBookMark As Variant      'marcador del formulario
Dim mbNotFound As Integer      'utilizado por la función de búsqueda
Dim mbEditFlag As Integer      'modo de edición
Dim mbAddNewFlag As Integer    'modo de agregar
Dim mbDataChanged As Integer   'indicador de modificación de datos
Dim mfrmFind As New frmFindForm      'instancia del formulario de búsqueda
Dim mlNumRows As Long          'total de filas del Recordset

Private Sub cmdAdd_Click()
  On Error GoTo AddErr

  'establece el modo
  mrsFormRecordset.AddNew
  lblStatus.Caption = MSG1
  mbAddNewFlag = True
  If mrsFormRecordset.RecordCount > 0 Then
    mvBookMark = mrsFormRecordset.Bookmark
  Else
    mvBookMark = vbNullString
  End If

  picChangeButtons.Visible = True
  picViewButtons.Visible = False
  hsclCurrRow.Enabled = False
  
  ClearDataFields Me, mrsFormRecordset.Fields.Count
  txtFieldData(0).SetFocus
  mbDataChanged = False
  Exit Sub

AddErr:
  ShowError
End Sub

Private Sub cmdCancel_Click()
   On Error Resume Next

   picChangeButtons.Visible = False
   picViewButtons.Visible = True
   hsclCurrRow.Enabled = True

   mbEditFlag = False
   mbAddNewFlag = False
   mrsFormRecordset.CancelUpdate
   DBEngine.Idle dbFreeLocks
   DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
   mbDataChanged = False

End Sub

Private Sub cmdMove_Click()
  On Error GoTo MVErr
  
  Dim sBookMark As String
  Dim sRows As String
  Dim lRows As Long
  
  sRows = InputBox(MSG2 & vbCrLf & MSG3)
  If Len(sRows) = 0 Then Exit Sub
  
  lRows = CLng(sRows)
  mrsFormRecordset.Move lRows
  
  'verifica que no se han sobrepasado los límites del Recordset
  If mrsFormRecordset.EOF Then
    mrsFormRecordset.MoveLast
  ElseIf mrsFormRecordset.BOF Then
    mrsFormRecordset.MoveFirst
  End If
  
  sBookMark = mrsFormRecordset.Bookmark  'guarda la nueva posición
  'ahora tenemos que volver a colocar la barra de desplazamiento para reflejar el movimiento
  If mlNumRows > 32767 Then
    hsclCurrRow.Value = (mrsFormRecordset.PercentPosition * 32767) / 100 + 1
  ElseIf mlNumRows > 99 Then
    hsclCurrRow.Value = (mrsFormRecordset.PercentPosition * mlNumRows) / 100 + 1
  Else
    hsclCurrRow.Value = mrsFormRecordset.PercentPosition
  End If
  mrsFormRecordset.Bookmark = sBookMark
  
  Exit Sub
  
MVErr:
  ShowError
End Sub

Private Sub hsclCurrRow_Change()
  On Error GoTo SCRErr
  
  Static nPrevVal As Integer
  Dim rsTmp As Recordset

  'comprueba si hay nuevas filas
  On Error Resume Next
  Set rsTmp = mrsFormRecordset.Clone()
  rsTmp.MoveNext
  If mrsFormRecordset.RecordCount > mlNumRows Then
    mlNumRows = mrsFormRecordset.RecordCount
    SetScrollBar
  End If
  On Error GoTo SCRErr

  'de acuerdo con el número de filas, la lógica de posicionamiento dentro
  'del Recordset es diferente
  If mlNumRows > 0 Then
    If mlNumRows > 99 Then   '32767
      'si hay más de 32767, tenemos que utilizar los métodos Move puesto que
      'la barra de desplazamiento tiene el límite en 32767, de modo que si no
      'aplicamos esta lógica, sería imposible llegar a todos los registros con un pequeño
      'cambio en la barra de desplazamiento
      If hsclCurrRow.Value - nPrevVal = 1 Then
        mrsFormRecordset.MoveNext
      ElseIf hsclCurrRow.Value - nPrevVal = -1 Then
        mrsFormRecordset.MovePrevious
      Else
        If mlNumRows > 32767 Then
          mrsFormRecordset.PercentPosition = (hsclCurrRow.Value / 32767) * 100 + 0.005
        Else
          mrsFormRecordset.PercentPosition = (hsclCurrRow.Value / mlNumRows) * 100 + 0.005
        End If
      End If
      nPrevVal = hsclCurrRow.Value
'    ElseIf mlNumRows > 99 Then
'      'hay que calcular la posición cuando hay más de 99 registros
'      mrsFormRecordset.PercentPosition = (hsclCurrRow.Value / mlNumRows) * 100 + 0.005
    Else
      mrsFormRecordset.PercentPosition = hsclCurrRow.Value
    End If
  End If
  DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
  mbDataChanged = False

  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  Exit Sub

SCRErr:
  ShowError
End Sub

Private Sub txtFieldData_Change(Index As Integer)
  'establece el indicador si los datos han cambiado
  'se restablece a False cuando se presenta un nuevo registro
  mbDataChanged = True
End Sub

Private Sub txtFieldData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = &H73 Then   'F4
    lblFieldName_DblClick Index

  ElseIf KeyCode = 34 And vsbScrollBar.Visible Then
    'página hacia abajo con más de 10 campos
    vsbScrollBar.Value = vsbScrollBar.Value - 3000

  ElseIf KeyCode = 33 And vsbScrollBar.Visible Then
    'página hacia arriba con más de 10 registros
    vsbScrollBar.Value = vsbScrollBar.Value + 3000

  End If
End Sub

Private Sub txtFieldData_KeyPress(Index As Integer, KeyAscii As Integer)
  'sólo permite volver en los modos de edición y de agregar
  If mbEditFlag Or mbAddNewFlag Then
    If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{Tab}"
    End If

  'ignora las teclas si no se está en los modos de edición o de agregar
  ElseIf mbEditFlag = False And mbAddNewFlag = False Then
    KeyAscii = 0
  End If

End Sub

Private Sub txtFieldData_LostFocus(Index As Integer)
  On Error GoTo FldDataErr

  If mbDataChanged Then
    'almacena los datos en el campo
    mrsFormRecordset(Index) = txtFieldData(Index)
  End If

  'lo restablece por válido o por condición de error
  mbDataChanged = False
  Exit Sub

FldDataErr:
  'lo restablece por válido o por condición de error
  mbDataChanged = False
  ShowError
End Sub

Private Sub lblFieldName_DblClick(Index As Integer)
  On Error GoTo ZoomErr

  If mrsFormRecordset(Index).Type = dbText Or mrsFormRecordset(Index).Type = dbMemo Then
     If mrsFormRecordset(Index).Type = dbText Then
       gsZoomData = txtFieldData(Index).Text
     ElseIf mrsFormRecordset(Index).FieldSize() < gnGETCHUNK_CUTOFF Then
       gsZoomData = txtFieldData(Index).Text
     Else
       'agrega el resto de los datos del campo con Getchunk
       MsgBar "Obteniendo datos de campos Memo", True
       Screen.MousePointer = vbHourglass
       gsZoomData = txtFieldData(Index).Text & _
         StripNonAscii(mrsFormRecordset(Index).GetChunk(gnGETCHUNK_CUTOFF, gnMAX_MEMO_SIZE))
       Screen.MousePointer = vbDefault
       MsgBar vbNullString, False
     End If
     frmZoom.Caption = Mid(lblFieldName(Index).Caption, 1, Len(lblFieldName(Index).Caption) - 1)
     If mbAddNewFlag Or mbEditFlag Then
       frmZoom.cmdSave.Visible = True
       frmZoom.cmdCloseNoSave.Visible = True
     Else
       frmZoom.cmdClose.Visible = True
     End If
     If mrsFormRecordset(Index).Type = dbText Then
       frmZoom.txtZoomData.Text = gsZoomData
       frmZoom.Height = 1125
     Else
       frmZoom.txtMemo.Text = gsZoomData
       frmZoom.txtMemo.Visible = True
       frmZoom.txtZoomData.Visible = False
       frmZoom.Height = 2205
     End If

     frmZoom.Show vbModal
     If (mbAddNewFlag Or mbEditFlag) And gsZoomData <> "__CANCELLED__" Then
       If mrsFormRecordset(Index).Type = dbText And Len(gsZoomData) > mrsFormRecordset(Index).Size Then
         Beep
         MsgBox MSG4, 48
         txtFieldData(Index).Text = Mid(gsZoomData, 1, mrsFormRecordset(Index).Size)
       Else
         txtFieldData(Index).Text = gsZoomData
       End If
       mrsFormRecordset(Index) = txtFieldData(Index).Text
       mbDataChanged = False
     End If
  End If
  Exit Sub

ZoomErr:
  ShowError
End Sub

Private Sub cmdClose_Click()
  DBEngine.Idle dbFreeLocks
  Unload Me
End Sub

Private Sub vsbScrollBar_Change()
  Dim nTop As Integer

  nTop = vsbScrollBar.Value
  If (nTop - 1080) Mod gnCTLARRAYHEIGHT = 0 Then
    picFields.Top = nTop
  Else
    picFields.Top = ((nTop - 1080) \ gnCTLARRAYHEIGHT) * gnCTLARRAYHEIGHT + 1080
  End If

End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DelRecErr

  If MsgBox(MSG5, vbYesNo + vbQuestion) = vbYes Then
    mrsFormRecordset.Delete
    If gbTransPending Then gbDBChanged = True
    If mrsFormRecordset.EOF = False Then
      'comprueba si puede pasar al registro siguiente
      mrsFormRecordset.MoveNext
      If mrsFormRecordset.EOF And (mrsFormRecordset.RecordCount > 0) Then
        'tiene que haber sido el último registro, así que no podemos mover hacia delante
        mrsFormRecordset.MoveLast
      End If
    End If
    mlNumRows = mlNumRows - 1
    SetScrollBar
    mlNumRows = mrsFormRecordset.RecordCount
    DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
    mbDataChanged = False
  End If

  Exit Sub

DelRecErr:
  ShowError
End Sub

Private Sub cmdEdit_Click()
   On Error GoTo EditErr

  Dim nDelay As Long
  Dim nRetryCnt As Integer
  
  Screen.MousePointer = vbHourglass
RetryEdit:
   mrsFormRecordset.Edit
   lblStatus.Caption = MSG6
   mbEditFlag = True
   txtFieldData(0).SetFocus
   mvBookMark = mrsFormRecordset.Bookmark

   picChangeButtons.Visible = True
   picViewButtons.Visible = False
   hsclCurrRow.Enabled = False

   Screen.MousePointer = vbDefault
   Exit Sub

EditErr:
  If Err = 3260 And nRetryCnt < gnMURetryCnt Then
    nRetryCnt = nRetryCnt + 1
    DBEngine.Idle dbFreeLocks
    'Espera gnMUDelay segundos
    nDelay = Timer
    While Timer - nDelay < gnMUDelay
      'no hace nada
    Wend
    Resume RetryEdit
  Else
    ShowError
  End If
End Sub

Private Sub cmdFilter_Click()
  On Error GoTo FilterErr

  Dim sBookMark As String
  Dim recRecordset1 As Recordset, recRecordset2 As Recordset
  Dim sFilterStr As String

  If mrsFormRecordset.RecordCount = 0 Then Exit Sub

  sBookMark = mrsFormRecordset.Bookmark        'guarda el marcador
  Set recRecordset1 = mrsFormRecordset         'guarda el Recordset
  
  sFilterStr = InputBox(MSG7)
  If Len(sFilterStr) = 0 Then Exit Sub

  Screen.MousePointer = vbHourglass
  MsgBar MSG8, True
  mrsFormRecordset.Filter = sFilterStr
  Set recRecordset2 = mrsFormRecordset.OpenRecordset(mrsFormRecordset.Type) 'establece el filtro
  'fuerza la población para tener la cuenta de registros precisa
  recRecordset2.MoveLast
  recRecordset2.MoveFirst
  Set mrsFormRecordset = recRecordset2            'vuelve a asignar el objeto Recordset original

  'todo ha ido bien, así que se presenta el formulario con el primer registro
  mlNumRows = mrsFormRecordset.RecordCount
  SetScrollBar
  hsclCurrRow.Value = 0
  DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
  mbDataChanged = False
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  Exit Sub
  
FilterRecover:
  On Error Resume Next
  Set mrsFormRecordset = recRecordset1            'vuelve a asignar el original
  mrsFormRecordset.Bookmark = sBookMark           'vuelve al registro original
  Exit Sub

FilterErr:
  ShowError
  Resume FilterRecover
End Sub

Private Sub cmdFind_Click()
  On Error GoTo FindErr
  
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String

  'carga los nombres de las columnas en el formulario de búsqueda
  If mfrmFind.lstFields.ListCount = 0 Then
    For i = 0 To mrsFormRecordset.Fields.Count - 1
      mfrmFind.lstFields.AddItem Mid(lblFieldName(i).Caption, 1, Len(lblFieldName(i).Caption) - 1)
    Next
  End If

FindStart:

  'restablece los indicadores
  gbFindFailed = False
  gbFromTableView = False
  mbNotFound = False

  MsgBar MSG9, False
  mfrmFind.Show vbModal
  MsgBar MSG10, True
  If gbFindFailed Then    'búsqueda cancelada
    GoTo AfterWhile
  End If

  Screen.MousePointer = vbHourglass

  i = mfrmFind.lstFields.ListIndex
  sBookMark = mrsFormRecordset.Bookmark
  'busca el registro
  If mrsFormRecordset(i).Type = dbText Or mrsFormRecordset(i).Type = dbMemo Then
    sTmp = AddBrackets((mrsFormRecordset(i).Name)) & " " & gsFindOp & " '" & gsFindExpr & "'"
  Else
    sTmp = AddBrackets((mrsFormRecordset(i).Name)) + gsFindOp + gsFindExpr
  End If
  Select Case gnFindType
    Case 0
      mrsFormRecordset.FindFirst sTmp
    Case 1
      mrsFormRecordset.FindNext sTmp
    Case 2
      mrsFormRecordset.FindPrevious sTmp
    Case 3
      mrsFormRecordset.FindLast sTmp
  End Select
  mbNotFound = mrsFormRecordset.NoMatch

AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFailed Then    'vuelve a la fila original
    mrsFormRecordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox MSG11, 48
    mrsFormRecordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = mrsFormRecordset.Bookmark  'guarda la nueva posición
    'ahora tenemos que volver a colocar la barra de desplazamiento para que refleje el movimiento
    If mlNumRows > 99 Then
      hsclCurrRow.Value = (mrsFormRecordset.PercentPosition * mlNumRows) / 100 + 1
    Else
      hsclCurrRow.Value = mrsFormRecordset.PercentPosition
    End If
    mrsFormRecordset.Bookmark = sBookMark
  End If

  DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
  mbDataChanged = False

  MsgBar vbNullString, False
  Exit Sub

FindErr:
  Screen.MousePointer = vbDefault
  If Err <> gnEOF_ERR Then
    ShowError
  Else
    mbNotFound = True
    Resume Next
  End If
End Sub

Private Sub Form_Load()
  Dim sTmp As String             'cadena temporal del nombre del Recordset
  Dim nFieldType As Integer      'tipo del campo actual
  Dim i As Integer, j As Integer 'índices

  On Error GoTo DynasetErr
   
   
  cmdAdd.Caption = BUTTON1
  cmdEdit.Caption = BUTTON2
  cmdDelete.Caption = BUTTON3
  cmdClose.Caption = BUTTON4
  cmdSort.Caption = BUTTON5
  cmdFilter.Caption = BUTTON6
  cmdMove.Caption = BUTTON7
  cmdFind.Caption = BUTTON8
  cmdCancel.Caption = BUTTON9
  cmdUpdate.Caption = BUTTON10
  lblFieldHdr.Caption = Label1
  lblFieldValue.Caption = Label2
  
  'mrsFormRecordset es una variable pública a nivel de módulo
  'que tiene que ser establecida antes de mostrar este formulario
   
   
  'establece el tipo de bloqueo (coméntelo si se va a utilizar de forma independiente)
  If gsDataType = gsMSACCESS And mrsFormRecordset.Type <> dbOpenSnapshot Then
    mrsFormRecordset.LockEdits = gnMULocking
  End If

  'obtiene la cuenta de filas
  With mrsFormRecordset
    If .RecordCount > 0 Then
      'pasa al siguiente y después al anterior; así obtiene la cuenta de registros
      .MoveNext
      .MovePrevious
    End If
    mlNumRows = .RecordCount
  End With
  SetScrollBar

  'carga los controles en el formulario del Recordset
  lblFieldName(0).Visible = True
  txtFieldData(0).Visible = True
  nFieldType = mrsFormRecordset(0).Type
  txtFieldData(0).Width = GetFieldWidth(nFieldType)
  If nFieldType = dbText Then txtFieldData(0).MaxLength = mrsFormRecordset(0).Size
  txtFieldData(0).TabIndex = 0
  For i = 1 To mrsFormRecordset.Fields.Count - 1
    picFields.Height = picFields.Height + gnCTLARRAYHEIGHT
    Load lblFieldName(i)
    lblFieldName(i).Top = lblFieldName(i - 1).Top + gnCTLARRAYHEIGHT
    lblFieldName(i).Visible = True
    Load txtFieldData(i)
    txtFieldData(i).Top = txtFieldData(i - 1).Top + gnCTLARRAYHEIGHT
    txtFieldData(i).Visible = True
    nFieldType = mrsFormRecordset.Fields(i).Type
    txtFieldData(i).Width = GetFieldWidth(nFieldType)
    If nFieldType = dbText Then txtFieldData(i).MaxLength = mrsFormRecordset(i).Size
    txtFieldData(i).TabIndex = i
  Next

  'cambia el tamaño de la ventana principal
  Me.Width = 5580
  If i <= 10 Then
    Me.Height = ((i + 1) * gnCTLARRAYHEIGHT) + 1600
  Else
    Me.Height = 4368
    Me.Width = Me.Width + 260
    vsbScrollBar.Visible = True
    vsbScrollBar.Min = 1080
    vsbScrollBar.Max = 1080 - (i * gnCTLARRAYHEIGHT) + 2240
  End If

  'presenta los nombres de los campos
  For i = 0 To mrsFormRecordset.Fields.Count - 1
    lblFieldName(i).Caption = mrsFormRecordset(i).Name & ":"
  Next

  DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
  mbDataChanged = False

  Me.Left = 1000
  Me.Top = 1000
  
  MsgBar vbNullString, False
  Exit Sub

DynasetErr:
  ShowError
  Unload Me
End Sub

Private Sub Form_Resize()
  On Error Resume Next

  Dim nHeight As Integer
  Dim i As Integer
  Dim nTotWidth As Integer
  Const nHeightFactor = 1420

  If WindowState <> 1 Then   'no minimizado
    MsgBar MSG12, True
    'asegura que el formulario está alineado con un campo
    nHeight = Height
    If (nHeight - nHeightFactor) Mod gnCTLARRAYHEIGHT <> 0 Then
      Me.Height = ((nHeight - nHeightFactor) \ gnCTLARRAYHEIGHT) * gnCTLARRAYHEIGHT + nHeightFactor
    End If
    'cambia el tamaño de la barra de estado
    picMoveButtons.Top = Me.Height - 650
    'cambia el tamaño de la barra de desplazamiento
    vsbScrollBar.Height = picMoveButtons.Top - (picViewButtons.Top - picFldHdr.Height) - 1320
    vsbScrollBar.Left = Me.Width - 360
    If mrsFormRecordset.Fields.Count > 10 Then
      picFields.Width = Me.Width - 260
      nTotWidth = vsbScrollBar.Left - 20
    Else
      picFields.Width = Me.Width - 20
      nTotWidth = Me.Width - 50
    End If
    picFldHdr.Width = Me.Width - 20
    'amplía los campos si es posible
    For i = 0 To mrsFormRecordset.Fields.Count - 1
      lblFieldName(i).Width = 0.3 * nTotWidth
      txtFieldData(i).Left = lblFieldName(i).Width + 20
      If mrsFormRecordset(i).Type = dbText Or mrsFormRecordset(i).Type = dbMemo Then
        txtFieldData(i).Width = 0.7 * nTotWidth - 250
      End If
    Next
    lblFieldValue.Left = txtFieldData(0).Left
    hsclCurrRow.Width = picMoveButtons.Width \ 2
    lblStatus.Width = picMoveButtons.Width \ 2
    lblStatus.Left = hsclCurrRow.Width + 10
  End If
  MsgBar vbNullString, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next

  Unload mfrmFind   'se deshace del formulario de búsqueda
  mrsFormRecordset.Close          'cierra el formulario
  DBEngine.Idle dbFreeLocks
  MsgBar vbNullString, False
End Sub

Private Sub cmdSort_Click()
  On Error GoTo SortErr

  Dim sBookMark As String
  Dim recRecordset1 As Recordset, recRecordset2 As Recordset
  Dim SortStr As String

  If mrsFormRecordset.RecordCount = 0 Then Exit Sub

  sBookMark = mrsFormRecordset.Bookmark        'guarda el marcador
  Set recRecordset1 = mrsFormRecordset            'guarda el Recordset
  
  SortStr = InputBox(MSG13)
  If Len(SortStr) = 0 Then Exit Sub

  Screen.MousePointer = vbHourglass
  MsgBar MSG14, True
  mrsFormRecordset.Sort = SortStr
  'establece el orden
  Set recRecordset2 = mrsFormRecordset.OpenRecordset(mrsFormRecordset.Type)
  Set mrsFormRecordset = recRecordset2            'vuelve a asignar el objeto Recordset original

  'todo ha ido bien, así que presenta el formulario con el primer registro
  mlNumRows = mrsFormRecordset.RecordCount
  hsclCurrRow.Value = 0
  DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
  mbDataChanged = False
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  Exit Sub

SortRecover:
  On Error Resume Next
  Set mrsFormRecordset = recRecordset1           'vuelve a asignar el original
  mrsFormRecordset.Bookmark = sBookMark          'vuelve al registro original
  Exit Sub

SortErr:
  ShowError
  Resume SortRecover
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  Dim nDelay As Long
  Dim nRetryCnt As Integer

  Screen.MousePointer = vbHourglass
RetryUpd:
  mrsFormRecordset.Update
  If gbTransPending Then gbDBChanged = True

  If mbAddNewFlag Then
    mlNumRows = mlNumRows + 1
    SetScrollBar
    'pasa al registro nuevo
    mrsFormRecordset.Bookmark = mrsFormRecordset.LastModified
  End If

  picChangeButtons.Visible = False
  picViewButtons.Visible = True
  hsclCurrRow.Enabled = True
  mbEditFlag = False
  mbAddNewFlag = False
  DisplayCurrentRecord Me, mrsFormRecordset, mlNumRows, mbAddNewFlag
  mbDataChanged = False
  DBEngine.Idle dbFreeLocks

  Screen.MousePointer = vbDefault
  Exit Sub

UpdateErr:
  'comprueba si hay un error de bloqueo
  If Err = 3260 And nRetryCnt < gnMURetryCnt Then
    nRetryCnt = nRetryCnt + 1
    mrsFormRecordset.Bookmark = mrsFormRecordset.Bookmark   'Cancela la actualización
    DBEngine.Idle dbFreeLocks
    nDelay = Timer
    'Espera gnMUDelay segundos
    While Timer - nDelay < gnMUDelay
      'no hace nada
    Wend
    Resume RetryUpd
  Else
    ShowError
  End If
End Sub

Private Sub SetScrollBar()
  On Error Resume Next
  
  If mlNumRows < 2 Then
    hsclCurrRow.Max = 100
    hsclCurrRow.SmallChange = 1 '00
    hsclCurrRow.LargeChange = 100
  ElseIf mlNumRows > 32767 Then
    hsclCurrRow.Max = 32767
    hsclCurrRow.SmallChange = 1
    hsclCurrRow.LargeChange = 1000
  ElseIf mlNumRows > 99 Then
    hsclCurrRow.Max = mlNumRows
    hsclCurrRow.SmallChange = 1
    hsclCurrRow.LargeChange = mlNumRows \ 20
  Else
    'tiene que estar entre 2 y 100
    hsclCurrRow.Max = 100
    hsclCurrRow.SmallChange = 100 \ (mlNumRows - 1)
    hsclCurrRow.LargeChange = (100 \ (mlNumRows - 1)) * 10
  End If
  'sale y vuelve a entrar para arreglar el dibujo de la barra
  txtFieldData(0).SetFocus
  hsclCurrRow.SetFocus
End Sub

