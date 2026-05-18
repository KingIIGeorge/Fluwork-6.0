Attribute VB_Name = "modMain"
Option Explicit


Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal msg&, ByVal wp&, ByVal lp&)
Private Declare Sub SetFocus Lib "user32" (ByVal hwnd&)
Private Declare Function GetParent Lib "user32" (ByVal hwnd&) As Long
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const VK_F = 70  ' VK_A a VK_Z coinciden con sus equivalentes ASCII : 'A' a'Z'
Dim hwndMenu       As Long       'necesario para pasar las pulsaciones de menú a VB

Global gVBInstance  As VBIDE.VBE       'instancia del IDE de VB
Global gwinWindow   As VBIDE.Window    'se usa para asegurarse de que sólo se ejecuta una instancia
Global gdocTabOrder As Object          'objeto documento de usuario


Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"

Sub AddToINI()
  'Este código agrega una entrada a VB5.INI
  'Debe ejecutarse desde la ventana Inmediato 
  Debug.Print WritePrivateProfileString("Add-Ins32", "TabOrder.Connect", "0", "vbaddin.ini")
End Sub

Function InRunMode(VBInst As VBIDE.VBE) As Boolean
  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function

Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
  If Shift <> 4 Then Exit Sub
  If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
  If gVBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub
  
  If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.hwnd)
  PostMessage hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
  KeyCode = 0
  SetFocus hwndMenu
End Sub

Function FindHwndMenu&(ByVal hwnd&)
  Dim h As Long
  
Loop2:
  h = GetParent(hwnd)
  If h = 0 Then FindHwndMenu = hwnd: Exit Function
  hwnd = h
  GoTo Loop2
End Function

