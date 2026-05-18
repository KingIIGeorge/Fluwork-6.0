VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoffeeNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' El módulo ICoffeeNotify determina una interfaz
'   que contiene el método de llamada de retorno,
'   CoffeeReady, y una propiedad utilizada para
'   almacenar el ID del cliente (que asigna
'   CoffeeMonitor2).
'

' CoffeeReady contiene en su implementación todo
' -----------   el código que desea ejecutar el
'   cliente cuando se produce una llamada de retorno.
'
Public Sub CoffeeReady()
End Sub

' NotifyID debe proporcionar en su implementación un
' --------    Long para almacenar el número de ID. 
'   CoffeeMonitor2 lo usará para buscar el objeto de
'   llamada de retorno cuando se solicite la desconexión.
'
Public Property Get NotifyID() As Long
End Property
Public Property Let NotifyID(ByVal NewValue As Long)
End Property
