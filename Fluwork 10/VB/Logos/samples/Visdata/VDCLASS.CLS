VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VisData Database Utility"
'>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>

Public VBInstance As VBIDE.VBE

Implements IDTExtensibility

'---------------------------------------------------
'este método se utiliza como llamada desde un programa
'externo o desde el IDE de VB
'---------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
  On Error GoTo LVDErr
    
  Dim rc As Long
  Set gVDClass = Me
  'esto establece el controlador de la instancia de VB
  'utilizada por el formulario DataFormDesigner
  Set VBInstance = VBInst
  frmMDI.mnuUDataFormDesigner.Visible = True
  frmMDI.WindowState = vbNormal
  
  frmMDI.SetWindowParent
  
  Exit Sub
    
LVDErr:
  MsgBox Err.Description
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
  On Error Resume Next
  Unload frmMDI
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'---------------------------------------------------
'este método puede ser llamado desde
'cualquier aplicación mediante Automatización OLE como en:
'Dim x As Object
'Set x = CreateObject("VisData.VisDataClass")
'x.VDOpenDatabase "c:\vb\biblio.mdb", "", False
'---------------------------------------------------
Sub VDOpenDatabase(sDatabaseName As String, sConnect As String, bReadOnly As Integer)
  On Error GoTo RVDErr
    
  frmMDI.mnuUDataFormDesigner.Visible = False
  
  gnReadOnly = bReadOnly
  If Len(sConnect) = 0 Then
    'tiene que ser una base de datos MDB de Microsoft Access
    gsDataType = gsMSACCESS
    gsDBName = sDatabaseName
    OpenLocalDB True
  ElseIf UCase(Left(sConnect, 5)) = gsSQLDB Then
    'tiene que ser una base de datos ODBC
    gsDataType = gsSQLDB
    'establece otras variables de ODBC
    GetODBCConnectParts sConnect
    gsDBName = sDatabaseName
    SendKeys "%AAO{Enter}"
  Else
    'tiene que ser una base de datos ISAM local
    gsDataType = sConnect
    gsDBName = sDatabaseName
    OpenLocalDB True
  End If
  
  Exit Sub
    
RVDErr:
  MsgBox Err.Description
End Sub

'--------------------------------------------------------
'este método simplemente cierra Visdata desde el cliente
'--------------------------------------------------------
Sub VDClose()
  Unload frmMDI
End Sub

