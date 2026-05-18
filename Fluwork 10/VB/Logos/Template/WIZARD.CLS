VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Wizard Template"
Option Explicit

Implements IDTExtensibility

Dim mcbMenuCommandBar         As Office.CommandBarControl  'objeto barra de comandos
Public WithEvents MenuHandler As CommandBarEvents          'controlador de eventos de la barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1
Dim mfrmWizard As frmWizard
Dim VBInstance As VBIDE.VBE


'------------------------------------------------------
'este método agrega el complemento al menú de VB
'lo llama el Administrador de complementos de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal LoadMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
   On Error GoTo error_handler
   
   Set VBInstance = VBInst
   
   If LoadMode = vbext_cm_External Then
       'Lo usa la barra de herramientas del Asistente para iniciar este Asistente
       LoadMe
   Else
       Set mcbMenuCommandBar = AddToAddInCommandBar(VBInst, LoadResString(15), LoadResPicture(5000, 0))
       'inicia el evento
       Set Me.MenuHandler = VBInst.Events.CommandBarEvents(mcbMenuCommandBar)
   End If
  
   Exit Sub
     
error_handler:
   MsgBox Err.Description
End Sub

'------------------------------------------------------
'este método elimina el complemento del menú de VB
'lo llama el Administrador de complementos de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    'elimina la entrada de la barra de comandos
    mcbMenuCommandBar.Delete
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
  'es necesario un stub
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
  'es necesario un stub
End Sub

'este evento se activa al hacer clic en el entorno
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    LoadMe
End Sub

Private Sub LoadMe()
    Set mfrmWizard = New frmWizard
    'pasa la instancia de VB al módulo del Asistente
    Set mfrmWizard.VBInst = VBInstance
    'carga y muestra el formulario
    mfrmWizard.Show vbModal
    Set mfrmWizard = Nothing
End Sub

