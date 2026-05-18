VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' La interfaz IDebug está implementada por todos
'   los objetos del proyecto. Contiene
'   métodos de depuración y
'   propiedades (en ProgWOb, sólo hay
'   una de ellas: la propiedad DebugID).
'
' ¿Por qué tener una interfaz separada?  En este
'   proyecto, la clase Polygon se utiliza de
'   dos maneras: los objetos Polygon se crean
'   y la interfaz de Polygon
'   está implementada por otras dos clases.
'   De este modo, si ponemos una propiedad DebugID en
'   la interfaz predeterminada de Polygon, tendríamos
'   que poner propiedades DebugID sin uso
'   en dichas implementaciones, igual
'   que se hace con el método TimeTest.
'
' El tener una interfaz separada sólo
'   requiere una línea de código más (la
'   instrucción Implements IDebug) que el poner
'   DebugID en la interfaz predeterminada
'   de cada clase.
'
'
' DebugID se implementa proporcionando un
'   Long para almacenar el ID de
'   depuración, y llamando a DebugInit para establecer
'   el ID (en el evento Initialize).
'
Public Property Get DebugID() As Long
End Property
