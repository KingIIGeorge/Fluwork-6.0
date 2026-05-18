Attribute VB_Name = "Module1"
' Constantes públicas
Public Const conMCIAppTitle = "Aplicación del control MCI"

' Estas constantes se definen en mmsystem.h.
Public Const conMCIErrInvalidDeviceID = 30257
Public Const conMCIErrDeviceOpen = 30263
Public Const conMCIErrCannotLoadDriver = 30266
Public Const conMCIErrUnsupportedFunction = 30274
Public Const conMCIErrInvalidFile = 30304

#If Win32 Then
    Declare Function GetFocus Lib "User32" () As Long
#Else
    Declare Function GetFocus Lib "User" () As Integer
#End If

' Variables públicas
Public DialogCaption As String

Public Sub DisplayWarning()
    Dim WMsg As String
    Dim WTitle As String
   
    WMsg = "Al insertar un CD de audio en Windows 95 o Windows NT, puede que se inicie un reproductor de CD de audio y comienze a reproducir el CD automáticamente. Si ocurre esto, seleccione el reproductor de CD predeterminado de la barra de tareas de Windows y ciérrelo antes de continuar."
    WTitle = "Advertencia"
    
    MsgBox WMsg, vbOKOnly, WTitle
End Sub

