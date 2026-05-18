VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' > Encontrará una descripción general de esta aplicación de
'   ejemplo buscando Coffee en la Ayuda en pantalla.
' > AboutCof.Txt, en la carpeta Documentos relacionados de
'   CoffWat2.vbp, también contiene información sobre el ejemplo.

' Clase Connector
' ---------------
' Connector permite que varios clientes compartan un solo
'   objeto CoffeeMonitor. La clase Connector tiene en su
'   propiedad Instancing el valor MultiUse, por lo que cada
'   cliente puede crear su propio Connector. Cada uno de
'   los objetos Connector devuelve una referencia al objeto
'   CoffeeMonitor único compartido, de modo que todos los
'   clientes comparten el mismo CoffeeMonitor. (Vea la
'   propiedad CoffeeMonitor, más abajo).
'
' Al leer "Creación de un componente EXE ActiveX", en los Manuales
'   en pantalla, observará una diferencia:  Esta versión de
'   Connector no utiliza una 'cuenta de uso' para determinar
'   cuándo hay que establecer Nothing en gCoffeeMonitor. No
'   es necesario código en el evento Initialize para incrementar
'   la cuenta de uso, ni código en el evento Terminate para
'   reducirla y comprobar su valor.
'
' Así se resuelve un error descrito en el tema "Uso de la
'   clase CoffeeMonitor compartida", en el capítulo mencionado
'   anteriormente. Si desea una explicación de por qué no
'   es necesaria la cuenta de uso, y de cómo se ha corregido
'   el error, vea el módulo de clase de CoffeeMonitor.


' La propiedad CoffeeMonitor devuelve siempre la referencia
' --------------------------  global única a la instancia
'   compartida de CoffeeMonitor (que se crea la primera vez
'   que un cliente solicita un CoffeeMonitor).
'
Public Property Get CoffeeMonitor() As CoffeeMonitor
    ' Si no se ha creado el objeto CoffeeMonitor compartido
    '   lo crea y almacena una referencia global al mismo.
    If gCoffeeMonitor Is Nothing Then
        Set gCoffeeMonitor = New CoffeeMonitor
    End If
    
    Set CoffeeMonitor = gCoffeeMonitor
End Property
    
