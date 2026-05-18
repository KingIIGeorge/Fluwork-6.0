VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VB TabOrder Window"
Option Explicit

Implements IDTExtensibility

Public WithEvents PrjHandler  As VBProjectsEvents         'controlador de eventos de proyectos
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler  As VBComponentsEvents       'controlador de eventos de componentes
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler  As VBControlsEvents         'controlador de eventos de controles
Attribute CtlHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents         'controlador de eventos de barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1
Dim mcbMenuCommandBar         As Office.CommandBarControl 'objeto barra de comandos

Const guidMYTOOL$ = "{B7AFC8D0-EBE5-11cf-A497-00A0C911E8B0}"

Public NonModalApp As Boolean                             'utilizado por la barra de herramientas Complementos
Attribute NonModalApp.VB_VarMemberFlags = "40"

Sub Show()
Attribute Show.VB_MemberFlags = "40"
  On Error GoTo ShowErr
  
  gwinWindow.Visible = True
  gdocTabOrder.RefreshList 3
  
  Exit Sub
ShowErr:
  MsgBox Err.Description
End Sub

Private Sub Class_Initialize()
  NonModalApp = True   'utilizado por la barra de herramientas Complementos
End Sub

'--------------------------------------------------------
' Este método agrega el complemento al menú Herramientas
' de VB. Es invocado por el Administrador de
' complementos de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
  On Error GoTo IDTExtensibility_OnConnectionErr
  
  Dim aiTmp As AddIn
  
  ' Guarda la instancia de VB
  Set gVBInstance = VBInst

  If Not gwinWindow Is Nothing Then
    ' Ya en ejecución, así que sólo lo muestra
    Show
    If ConnectMode = vbext_cm_AfterStartup Then
      ' Iniciado desde el Administrador de complementos
      AddToCommandBar
    End If
    Exit Sub
  End If
  
  ' Crea la ventana de herramientas
  If ConnectMode = vbext_cm_External Then
    ' Es necesario ver si ya está en ejecución
    On Error Resume Next
    Set aiTmp = gVBInstance.Addins("TabOrder.Connect")
    On Error GoTo IDTExtensibility_OnConnectionErr
    If aiTmp Is Nothing Then
      ' La aplicación no está en el archivo VBADDIN.INI, por lo que no
      ' pertenece a la colección, así que intentamos usar el primer
      ' complemento de la misma para poner a funcionar la aplicación.
      ' Si no hay ningún complemento, se producirá un error y la
      ' aplicación no se ejecutará.
       Set gwinWindow = gVBInstance.Windows.CreateToolWindow(gVBInstance.Addins(1), "TabOrder.docTabOrder", LoadResString(10), guidMYTOOL$, gdocTabOrder)
    Else
      If aiTmp.Connect = False Then
        Set gwinWindow = gVBInstance.Windows.CreateToolWindow(aiTmp, "TabOrder.docTabOrder", LoadResString(10), guidMYTOOL$, gdocTabOrder)
      End If
    End If
  Else
    ' Debe haber sido llamado desde el Administrador de complementos
    Set gwinWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "TabOrder.docTabOrder", LoadResString(10), guidMYTOOL$, gdocTabOrder)
  End If

  ' Cierra el controlador del proyecto, de los componentes y de los controles.
  Set Me.PrjHandler = gVBInstance.Events.VBProjectsEvents
  Set Me.CmpHandler = gVBInstance.Events.VBComponentsEvents(Nothing)
  Set Me.CtlHandler = gVBInstance.Events.VBControlsEvents(Nothing, Nothing)
  
  If ConnectMode = vbext_cm_External Then
    ' Iniciado desde la barra de herramientas de complementos
    Show
  ElseIf ConnectMode = vbext_cm_AfterStartup Then
    ' Iniciado desde el Administrador de complementos
    AddToCommandBar
  End If

  Exit Sub
  
IDTExtensibility_OnConnectionErr:
  MsgBox Err.Description
  
End Sub

'-------------------------------------------------------
' Este evento elimina el menú de barra de comandos
' cuando lo llama el Administrdor de complementos de VB
'-------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
  On Error GoTo IDTExtensibility_OnDisconnectionErr
  ' Elimina la entrada de la barra de comandos
  mcbMenuCommandBar.Delete
  
  ' Guarda el estado del formulario para la siguiente vez que se cargue VB
  If gwinWindow.Visible Then
    SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "1"
  Else
    SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "0"
  End If
  
  Set gwinWindow = Nothing
  
IDTExtensibility_OnDisconnectionErr:
  
End Sub

' Este evento se produce cuando el IDE se ha cargado por completo.
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
  AddToCommandBar
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
  ' Comentario necesario para evitar que VBA
  ' elimine este código al compilar
End Sub

' Este evento se produce al hacer clic en el control de barra de comandos en el IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Show
End Sub

' Este evento se produce al agregar un control al formulario actual en el IDE
Private Sub CtlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
  If gwinWindow.Visible Then
    gdocTabOrder.ControlAdded VBControl
  End If
End Sub

' Este evento se produce cuando se cambia un control del formulario actual en el IDE
Private Sub CtlHandler_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
  If gwinWindow.Visible Then
    gdocTabOrder.ControlRenamed VBControl, OldName, OldIndex
  End If
End Sub

' Este evento se produce cuando se elimina un control del formulario actual en el IDE
Private Sub CtlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
  If gwinWindow.Visible Then
    gdocTabOrder.ControlRemoved VBControl
  End If
End Sub

' Este evento se produce cuando se activa un formulario en el IDE
Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
  On Error GoTo CmpHandler_ItemActivatedErr
  If gwinWindow.Visible Then
    gdocTabOrder.RefreshList 0
  End If
CmpHandler_ItemActivatedErr:
End Sub

' Este evento se produce cuando se selecciona un formulario en la ventana de proyecto
Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
  CmpHandler_ItemActivated VBComponent
End Sub

Sub AddToCommandBar()
Attribute AddToCommandBar.VB_MemberFlags = "40"
  On Error GoTo AddToCommandBarErr
  
  ' Se asegura de que la barra de herramientas estándar esté visible
  gVBInstance.CommandBars(2).Visible = True
  
  ' Lo agrega a la barra de comandos
  ' La línea siguiente agregará el Administrador de orden
  ' de tabulación a la barra de herramientas estándar, a
  ' la derecha del botón Cuadro de herramientas.
  Set mcbMenuCommandBar = gVBInstance.CommandBars(2).Controls.Add(1, , , gVBInstance.CommandBars(2).Controls.Count)
  ' Establece el título
  mcbMenuCommandBar.Caption = LoadResString(200)
  ' Copia el icono en el portapapeles
  Clipboard.SetData LoadResPicture(1000, 0)
  ' Establece el icono del botón
  mcbMenuCommandBar.PasteFace
  
  ' Cierra el evento
  Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
  
  ' Restablece el último estado
  If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") = "1" Then
    ' Establece ésto para mostrar el formulario al conectar
    Me.Show
  End If
  
  Exit Sub
    
AddToCommandBarErr:
  MsgBox Err.Description
End Sub

Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
  ' Ésto se ocupa del usuario que elimine el único proyecto
  If gwinWindow.Visible Then
    gdocTabOrder.RefreshList 0
  End If
End Sub
