VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connector2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' > Encontrará una descripción general de esta aplicación
'   de ejemplo buscando Coffee en la Ayuda en pantalla.
' > AboutCof.Txt, en la carpeta Documentos relacionados de
'   CoffWat2.vbp, también contiene información sobre el ejemplo.

' La clase Connector2 Permite que varios clientes compartan una
' -------------------  sola instancia de CoffeeMonitor2. La clase
'   Connector2 tiene en su propiedad Instancing el valor
'   MultiUse, por lo que cada cliente puede crear su propio Connector2.
'   Todos los objetos Connector2 devuelven una referencia al mismo
'   objeto CoffeeMonitor2 compartido, de modo que todos los clientes
'   comparten el mismo CoffeeMonitor2. (Vea la propiedad CoffeeMonitor2,
'   más abajo).

' La propiedad CoffeeMonitor2 devuelve siempre la referencia global
' ---------------------------   única a la instancia compartida de
'   CoffeeMonitor2.
'
Public Property Get CoffeeMonitor2() As CoffeeMonitor2
    ' Si el objeto CoffeeMonitor2 compartido no ha sido creado,
    '   lo crea y almacena una referencia global al mismo.
    If gCoffeeMonitor2 Is Nothing Then
        Set gCoffeeMonitor2 = New CoffeeMonitor2
    End If
    
    Set CoffeeMonitor2 = gCoffeeMonitor2
End Property
    

