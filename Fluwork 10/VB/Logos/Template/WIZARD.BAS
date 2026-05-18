Attribute VB_Name = "modWizard"
Option Explicit

Global Const WIZARD_NAME = "WizardTemplate"

Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

'Comandos de WinHelp
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELP_QUIT = &H2              '  Finaliza la Ayuda
Public Const HELP_CONTENTS = &H3&         '  Muestra el índice de contenidos
Public Const HELP_CONTEXT = &H1           '  Muestra el tema de ulTopic
Public Const HELP_INDEX = &H3             '  Muestra el índice

Global Const APP_CATEGORY = "Wizards"

Global Const CONFIRM_KEY = "ConfirmScreen"
Global Const DONTSHOW_CONFIRM = "DontShow"


'--------------------------------------------------------------------------
'es necesario ejecutar este procedimiento en la ventana Inmediato
'agregará la entrada a VBADDIN.INI si no existe para que el complemento
'esté disponible la próxima vez que se inicie VB
'--------------------------------------------------------------------------
Sub AddToINI()
    Debug.Print WritePrivateProfileString("Add-Ins32", WIZARD_NAME & ".Wizard", "0", "VBADDIN.INI")
End Sub

Function GetResString(nRes As Integer) As String
    Dim sTmp As String
    Dim sRetStr As String
  
    Do
        sTmp = LoadResString(nRes)
        If Right(sTmp, 1) = "_" Then
            sRetStr = sRetStr + VBA.Left(sTmp, Len(sTmp) - 1)
        Else
            sRetStr = sRetStr + sTmp
        End If
        nRes = nRes + 1
    Loop Until Right(sTmp, 1) <> "_"
    GetResString = sRetStr
  
End Function

Function GetField(sBuffer As String, sSep As String) As String
    Dim p As Integer
    
    p = InStr(sBuffer & sSep, sSep)
    GetField = VBA.Left(sBuffer, p - 1)
    sBuffer = Mid(sBuffer, p + Len(sSep))
  
End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next
    
    Dim ctl As Control
    Dim obj As Object
    
    'establece el título del formulario
    If IsNumeric(frm.Tag) Then
        frm.Caption = LoadResString(CInt(frm.Tag))
    End If
    
    'establece los títulos de los controles usando la propiedad
    'Caption de los elementos de menú y la propiedad Tag
    'de los otros controles
    For Each ctl In frm.Controls
        If TypeName(ctl) = "Menu" Then
            If IsNumeric(ctl.Caption) Then
                If Err = 0 Then
                    ctl.Caption = LoadResString(CInt(ctl.Caption))
                Else
                    Err = 0
                End If
            End If
        ElseIf TypeName(ctl) = "TabStrip" Then
            For Each obj In ctl.Tabs
                If IsNumeric(obj.Tag) Then
                    obj.Caption = LoadResString(CInt(obj.Tag))
                End If
                'comprueba que sea una sugerencia
                If IsNumeric(obj.ToolTipText) Then
                    If Err = 0 Then
                        obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
                    Else
                        Err = 0
                    End If
                End If
            Next
        ElseIf TypeName(ctl) = "Toolbar" Then
            For Each obj In ctl.Buttons
                If IsNumeric(obj.Tag) Then
                    obj.ToolTipText = LoadResString(CInt(obj.Tag))
                End If
            Next
        ElseIf TypeName(ctl) = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                If IsNumeric(obj.Tag) Then
                    obj.Text = LoadResString(CInt(obj.Tag))
                End If
            Next
        Else
            If IsNumeric(ctl.Tag) Then
                If Err = 0 Then
                    ctl.Caption = LoadResString(CInt(ctl.Tag))
                Else
                    Err = 0
                End If
            End If
            'comprueba que sea una sugerencia
            If IsNumeric(ctl.ToolTipText) Then
                If Err = 0 Then
                    ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
                Else
                    Err = 0
                End If
            End If
        End If
    Next

End Sub

'================================================================
'Objetivo:Reemplazar las cadenas <TOPIC_TEXT> de las cadenas
'         de archivo de recursos en la ubicación correcta
'         de los elementos localizados
'
'Entrada :sString = Cadena para buscar y reemplazar
'         sReplacement = Cadena para reemplazar elemento
'         sReplacement2 = Segunda cadena para reemplazar elemento
'
'Salida  :Nueva cadena con el elemento a reemplazar
'================================================================
Function ReplaceTopicTokens(sString As String, _
                            sReplacement As String, _
                            sReplacement2 As String) As String
    On Error Resume Next
    
    Dim p As Integer
    Dim sTmp As String
    
    Const TOPIC_TEXT = "<TOPIC_TEXT>"
    Const TOPIC_TEXT2 = "<TOPIC_TEXT2>"
    
    sTmp = sString
    Do
        p = InStr(sTmp, TOPIC_TEXT)
        If p Then
            sTmp = VBA.Left(sTmp, p - 1) + sReplacement + Mid(sTmp, p + Len(TOPIC_TEXT))
        End If
    Loop While p
    
    If Len(sReplacement2) > 0 Then
        Do
            p = InStr(sTmp, TOPIC_TEXT2)
            If p Then
                sTmp = VBA.Left(sTmp, p - 1) + sReplacement2 + Mid(sTmp, p + Len(TOPIC_TEXT2))
            End If
        Loop While p
    End If
    
    ReplaceTopicTokens = sTmp
  
End Function

Public Function GetResData(sResName As String, sResType As String) As String
    Dim sTemp As String
    Dim p As Integer
  
    sTemp = StrConv(LoadResData(sResName, sResType), vbUnicode)
    p = InStr(sTemp, vbNullChar)
    If p Then sTemp = VBA.Left$(sTemp, p - 1)
    GetResData = sTemp
End Function

Function AddToAddInCommandBar(VBInst As Object, sCaption As String, oBitmap As Object) As Object   'Office.CommandBarControl
    On Error GoTo AddToAddInCommandBarErr
    
    Dim c As Integer
    Dim cbMenuCommandBar As Object   'Office.CommandBarControl  'Objeto barra de comandos
    Dim cbMenu As Object
    
    'comprueba si podemos encontrar el menú Complementos
    Set cbMenu = VBInst.CommandBars(1).Controls(9)
    If cbMenu Is Nothing Then
        'no está disponible y salimos
        Exit Function
    End If
    
    'lo abrega a la barra de comandos
    Set cbMenuCommandBar = cbMenu.CommandBar.Controls.Add(1)
    c = cbMenu.CommandBar.Controls.Count - 1
    If cbMenu.CommandBar.Controls(c).BeginGroup And _
        Not cbMenu.CommandBar.Controls(c - 1).BeginGroup Then
        'comprueba si es el primer complemento y si se necesita una barra separadora
        cbMenuCommandBar.BeginGroup = True
    End If
    'establece el título
    cbMenuCommandBar.Caption = sCaption
    'no está realizado: establece OnAction (es necesario en este punto)
    cbMenuCommandBar.OnAction = "hello"
    'copia el icono al Portapapeles
    Clipboard.SetData oBitmap
    'establece el icono del botón
    cbMenuCommandBar.PasteFace
  
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
AddToAddInCommandBarErr:
  
End Function



