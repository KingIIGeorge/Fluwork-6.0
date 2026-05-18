VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDataGrid 
   ClientHeight    =   4605
   ClientLeft      =   3780
   ClientTop       =   2610
   ClientWidth     =   6180
   HelpContextID   =   2016191
   Icon            =   "DATAGRID.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Tag             =   "Recordset"
   Begin MSDBGrid.DBGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "DATAGRID.frx":014A
      Height          =   3795
      Left            =   0
      OleObjectBlob   =   "DATAGRID.frx":015F
      TabIndex        =   5
      Top             =   330
      Width           =   6180
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
      Top             =   4260
      Width           =   6180
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6180
      TabIndex        =   0
      Top             =   0
      Width           =   6180
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Actualizar"
         Height          =   330
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filtrar"
         Height          =   330
         Left            =   2880
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Ordenar"
         Height          =   330
         Left            =   1440
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   330
         Left            =   4320
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Top             =   0
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const BUTTON1 = "&Actualizar"
Const BUTTON2 = "&Ordenar"
Const BUTTON3 = "&Filtrar"
Const BUTTON4 = "&Cerrar"
Const DATACTL = "Haga clic con el botón secundario del mouse para ver las propiedades del control Data"
Const MSG1 = "Imposible filtrar un Recordset de tipo Table"
Const MSG2 = "Escriba la expresión de filtro:"
Const MSG3 = "Estableciendo un nuevo filtro"
Const MSG4 = "Imposible ordenar un Recordset de tipo Table"
Const MSG5 = "Introduzca la columna de orden:"
Const MSG6 = "Estableciendo la nueva columna de orden"
Const MSG7 = "¿Desea realizar los cambios?"
Const MSG8 = "¿Desea eliminar la fila actual?"
'>>>>>>>>>>>>>>>>>>>>>>>>


Public mrsFormRecordset As Recordset

Dim msSortCol As String
Dim mbCtrlKey As Integer


Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFilter_Click()
  On Error GoTo FilterErr

  Dim recRecordset1 As Recordset, recRecordset2 As Recordset
  Dim sFilterStr As String

  If datDataCtl.RecordsetType = vbRSTypeTable Then
    Beep
    MsgBox MSG1, 48
    Exit Sub
  End If
  
  Set recRecordset1 = datDataCtl.Recordset       'copia el Recordset
  
  sFilterStr = InputBox(MSG2)
  If Len(sFilterStr) = 0 Then Exit Sub

  Screen.MousePointer = 11
  MsgBar MSG3, True
  recRecordset1.Filter = sFilterStr
  Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type) 'establece el filtro
  Set datDataCtl.Recordset = recRecordset2       'vuelve a asignar el objeto Recordset original

  Screen.MousePointer = 0
  MsgBar vbNullString, False
  Exit Sub

FilterErr:
  Screen.MousePointer = 0
  MsgBar vbNullString, False
  MsgBox "Error:" & Err & " " & Error$
  Exit Sub

End Sub

Private Sub cmdRefresh_Click()
  On Error GoTo RefErr
  
  datDataCtl.Recordset.Requery
  
  Exit Sub
  
RefErr:
  MsgBox "Error:" & Err & " " & Error$
  Exit Sub
  
End Sub


Private Sub cmdSort_Click()
  On Error GoTo SortErr

  Dim recRecordset1 As Recordset, recRecordset2 As Recordset
  Dim SortStr As String

  If datDataCtl.RecordsetType = vbRSTypeTable Then
    Beep
    MsgBox MSG4, 48
    Exit Sub
  End If

  Set recRecordset1 = datDataCtl.Recordset        'copia el Recordset
  
  If Len(msSortCol) = 0 Then
    SortStr = InputBox(MSG5)
    If Len(SortStr) = 0 Then Exit Sub
  Else
    SortStr = msSortCol
  End If

  Screen.MousePointer = 11
  MsgBar MSG6, True
  recRecordset1.Sort = SortStr
  
  'establece el orden
  Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
  Set datDataCtl.Recordset = recRecordset2
  
  Screen.MousePointer = 0
  MsgBar vbNullString, False
  Exit Sub

SortErr:
  Screen.MousePointer = 0
  MsgBar vbNullString, False
  MsgBox "Error:" & Err & " " & Error$
  Exit Sub

End Sub

Private Sub datDataCtl_Validate(Action As Integer, Save As Integer)
  If Save Then
    If MsgBox(MSG7, vbYesNo + vbQuestion) <> vbYes Then
      'deshace los cambios
      datDataCtl.UpdateControls
    End If
  End If
End Sub

Private Sub Form_Load()
  On Error GoTo LoadErr
  
  cmdRefresh.Caption = BUTTON1
  cmdSort.Caption = BUTTON2
  cmdFilter.Caption = BUTTON3
  cmdClose.Caption = BUTTON4
  datDataCtl.Caption = DATACTL
  
  'mrsFormRecordset es una variable pública a nivel de módulo
  'que tiene que ser establecida antes de mostrar este formulario
  Set datDataCtl.Recordset = mrsFormRecordset
  
  Me.Width = 5865
  Me.Height = 5070
  
  Exit Sub

LoadErr:
  MsgBox "Error:" & Err & " " & Error$
  Unload Me

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.WindowState <> vbMinimized Then
    grdDataGrid.Height = Me.ScaleHeight - (picButtons.Height + datDataCtl.Height + 135)
  End If
End Sub

'----------------------------------------------------------
'esta subrutina carga la hoja de propiedades del control Data
'coméntela si va a hacer un uso independiente
'----------------------------------------------------------
Sub datDataCtl_MouseUp(BUTTON As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo DCPErr

  Dim i As Integer
  Dim recClone As Recordset
  Dim sTmpRS As String
  Dim sTmpDB As String
  
  If BUTTON = 2 Then
    Screen.MousePointer = 11
    sTmpRS = datDataCtl.RecordSource
    sTmpDB = datDataCtl.DatabaseName
    Set gDataCtlObj = datDataCtl
    frmDataCtlProp.Show vbModal
    If Not gDataCtlObj Is Nothing Then
      'verifica si el Recordset o la base de datos ha cambiado,
      'y borra los campos asociados si alguno ha cambiado
      datDataCtl.Refresh
      If sTmpRS <> gDataCtlObj.RecordSource Or sTmpDB <> gDataCtlObj.DatabaseName Then
        'restablece la cuadrícula de datos
        grdDataGrid.ReBind
      End If
      gbSettingDataCtl = False
    End If
    
  End If
  Exit Sub
  
DCPErr:
  MsgBox "Error:" & Err & " " & Error$
  Unload Me
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
  If MsgBox(MSG8, vbYesNo + vbQuestion) <> vbYes Then
    Cancel = True
  End If
End Sub


Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
  'ordenemos esta columna
  If datDataCtl.RecordsetType = vbRSTypeTable Then Exit Sub
  
  If datDataCtl.Recordset(ColIndex).Type > dbText Then Exit Sub
  
  'comprueba el uso de la tecla Ctrl para tener orden descendente
  If mbCtrlKey Then
    msSortCol = "[" & datDataCtl.Recordset(ColIndex).Name & "] desc"
    mbCtrlKey = 0 'lo restablece
  Else
    msSortCol = "[" & datDataCtl.Recordset(ColIndex).Name & "]"
  End If
  cmdSort_Click
  msSortCol = vbNullString 'lo restablece
  
End Sub

Private Sub grdDataGrid_MouseUp(BUTTON As Integer, Shift As Integer, x As Single, y As Single)
  mbCtrlKey = Shift
End Sub

