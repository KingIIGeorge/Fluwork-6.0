Attribute VB_Name = "Module1"
Option Explicit

Declare Function timeGetTime Lib "winmm.dll" () As Long

' CoffeeWatch debe ser un ActiveX exe, ya que
'   expone el objeto público NotifyMe empleado
'   en la demostración del método de llamada de
'   retorno. Nunca se inicia como servidor, sólo
'   como aplicación independiente.
Sub Main()
    If App.StartMode = vbSModeStandalone Then
        Form1.Show
    End If
End Sub
