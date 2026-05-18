Attribute VB_Name = "Module1"
Option Explicit
' > Encontrará una descripción general de esta aplicación
'   de ejemplo buscando Coffee en la Ayuda en pantalla.
' > AboutCof.Txt, en la carpeta Documentos relacionados de
'   CoffWat2.vbp, también contiene información sobre el ejemplo.

' gCoffeeMonitor contiene una referencia al objeto
' --------------  CoffeeMonitor compartido. Cuando un cliente
'   accede a la propiedad CoffeeMonitor de un objeto Connector,
'   la propiedad Get devuelve esta referencia.
'   (Vea el módulo de clase Connector).

Public gCoffeeMonitor As CoffeeMonitor
' gCoffeeMonitor2 contiene una referencia al objeto
' ---------------  CoffeeMonitor2 compartido. Cuando un cliente
'   accede a la propiedad CoffeeMonitor2 de un objeto Connector2,
'   la propiedad Get devuelve esta referencia.
'   (Vea el módulo de clase Connector2).
Public gCoffeeMonitor2 As CoffeeMonitor2

