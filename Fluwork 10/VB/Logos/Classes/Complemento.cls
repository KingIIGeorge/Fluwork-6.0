VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "My Add-In"
Option Explicit

Implements IDTExtensibility

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'controlador de eventos de la barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1


Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
End Sub

'------------------------------------------------------
'este método agrega el complemento a VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
    On Error GoTo error_handler
    
    'guarda la instancia de VB
    Set VBInstance = VBInst
    
    'éste es un buen sitio para colocar un punto de interrupción y
    'comprobar los objetos, propiedades y métodos del complemento
    Debug.Print VBInst.FullName

    If ConnectMode = vbext_cm_External Then
        'Lo usa el asistente de barras de herramientas para iniciar este asistente
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("My AddIn")
        'inicia el evento
        Set Me.MenuHandler = VBInst.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = vbext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'lo establece para mostrar el formulario de conexión
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'este método elimina el complemento de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    
    On Error Resume Next
    
    'elimina la entrada en la barra de comandos
    mcbMenuCommandBar.Delete
    
    'desactiva el complemento
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'lo establece para mostrar el formulario de conexión
        Me.Show
    End If
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'este método se activa al hacer clic en el menú del entorno
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'objeto barra de comandos
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'comprueba si se puede ver el menú Complementos
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'si no está disponible se sale
        Exit Function
    End If
    
    'lo agrega a la barra de comandos
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'establece el título
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function


