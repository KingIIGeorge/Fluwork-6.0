VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Triangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' >> Se ve mejor en la vista de Módulo completo. <<
'
' Almacenamiento para el número de ID de depuración.
Private mlngDebugID As Long
Implements IDebug

' La clase Triangle implementa la interfaz IShape
'   y la interfaz Polygon.
'   Se mantiene un Polygon interno,
'   y muchas actividades son delegadas
'   (dibujo, almacenamiento de puntos, etc.).
'   La propia interfaz de la clase Triangle
'   sólo tiene un método.

Implements IShape
Implements Polygon

' El objeto Polygon interno almacena realmente
'   los datos y hace la mayor parte del trabajo.
'   El Triangle mantiene referencias a la interfaz
'   Polygon y a la interfaz IShape
'   del Polygon interno.
Private mpyg As Polygon
Private mish As IShape

' -------------------------------------
' Esto marca el principio de la
'   implementación de la interfaz IShape.

' Se llama a IShape.DrawToPictureBox para
'   dibujar una forma,
'   de forma que cada clase tiene que suministrar
'   su propia implementación.
'
Private Sub IShape_DrawToPictureBox(ByVal pb As PictureBox)
    ' Delega en la interfaz IShape del
    '   Polygon interno.
    Call mish.DrawToPictureBox(pb)
End Sub

' El método IShape.TimeTest se utiliza para mostrar
'   el menor coste de las llamadas
'   a un método a través de una
'   interfaz implementada por varias
'   clases, en oposición con las llamadas a un
'   método similar a través de las interfaces
'   predeterminadas de las clases.
'
Private Sub IShape_TimeTest()
End Sub

' -------------------------------------
' Esto marca el principio de la
'   implementación de la interfaz
'   Polygon.

' Polygon.Color delega en el objeto
' Polygon interno.
'
Private Property Get Polygon_Color() As Long
    Polygon_Color = mpyg.Color
End Property
'
Private Property Let Polygon_Color(ByVal RHS As Long)
    mpyg.Color = RHS
End Property

' Polygon.TimeTest - Triangle tiene tres
'   métodos TimeTest, uno
'   en la interfaz IShape (utilizado para mostrar
'   el polimorfismo y la asociación previa), otro en su
'   propia interfaz (utilizado para mostrar la asociación
'   tardía), y éste. Éste es un efecto secundario
'   del hecho de que Rectangle
'   implemente la interfaz Polygon; no se
'   utiliza para nada.
Private Sub Polygon_TimeTest()
End Sub

' Polygon.GetPoint
'
Private Sub Polygon_GetPoint(ByVal intPoint As Integer, X As Single, Y As Single)
    ' Delega en el Polygon interno.
    Call mpyg.GetPoint(intPoint, X, Y)
End Sub

' Polygon.GetPointCount
'
Private Property Get Polygon_GetPointCount() As Integer
    ' No hay razón para delegar en el
    '   Polygon interno, un triángulo
    '   siempre tiene tres puntos.
    Polygon_GetPointCount = 3
End Property

' Polygon.SetPoints - Al implementar el
'   método SetPoints
'   de la interfaz Polygon, Triangle
'   ejecuta su propio código para verificar que
'   la matriz de entrada sólo contiene tres
'   puntos (seis Singles), y después
'   delega en el objeto Polygon
'   interno.
Private Sub Polygon_SetPoints(asngPoints() As Single)
    Dim blnBadArray As Boolean
    On Error Resume Next
    If UBound(asngPoints) <> 5 Then blnBadArray = True
    If Err.Number <> 0 Then blnBadArray = True
    If blnBadArray Then
        Err.Raise vbObjectError + 2082, , _
            "Un triángulo se especifica con una matriz basada en cero con seis valores, los valores pares (comenzando con 0) son valores de x y los valores impares son valores de y."
        Exit Sub
    End If
    ' Delega en el Polygon interno, que
    '   completa la validación de la matriz
    '   y la guarda.
    Call mpyg.SetPoints(asngPoints)
End Sub

' --------------------------------------
' Este es el principio de la propia
'    interfaz (predeterminada) de Triangle.

' El método TimeTest no tiene argumentos,
'   no hace nada y vuelve
'   inmediatamente. Se utiliza para
'   ilustrar la asociación tardía, en oposición
'   a la asociación previa ofrecida por
'   las llamadas a TimeTest en la interfaz
'   IShape.
Public Sub TimeTest()
End Sub

' --------------------------------------
' Este es el principio de los procedimientos
'   privados de Triangle (funciones de ayuda
'   y procedimientos de evento).

Private Sub Class_Initialize()
    Dim asngPoints(0 To 5) As Single
    ' Código de depuración.
    mlngDebugID = DebugInit(Me)
    '
    ' Crea el objeto Polygon interno y
    '   obtiene una referencia a su interfaz
    '   IShape.
    Set mpyg = New Polygon
    Set mish = mpyg
    ' Inicializa el Polygon interno.
    Call mpyg.SetPoints(asngPoints)
End Sub

Private Sub Class_Terminate()
    DebugTerm Me
End Sub

' -------- Implementación de IDebug --------
'
' IDebug.DebugID ofrece la manera de
'   distinguir los objetos. Es
'   requerido por los procedimientos de depuración
'   DebugInit, DebugTerm y DebugShow
'   declarados en modFriend.
'
Private Property Get IDebug_DebugID() As Long
    IDebug_DebugID = mlngDebugID
End Property


