VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rectangle"
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

' Rectangle implementa dos interfaces;
'   también tiene una propiedad (Filled) y
'   un método (TimeTest) de su propia
'   interfaz.

Implements IShape
Implements Polygon

' El objeto Polygon interno almacena
'   los datos y realiza algo de trabajo.
'   Rectangle mantiene referencias a la
'   interfaz de Polygon y a la interfaz de IShape
'   del objeto Polygon interno.
Private mpyg As Polygon
Private mish As IShape

' Almacenamiento para la propiedad Color (Implementación
'   de la interfaz de Polygon).
Private mrgbColor As Long

' Almacenamiento para la propiedad Filled (en la
'   interfaz predeterminada del objeto Rectangle).
Private mblnFilled As Boolean

' ---------------------------------------------------------
' Este es el principio de la implementación por parte de
'   Rectangle de la interfaz de
'   IShape.

' Se llama a IShape.DrawToPictureBox para
'   dibujar un objeto Shape,
'   de modo que cada clase Shape tiene que aportar
'   su propia implementación.
'
Private Sub IShape_DrawToPictureBox(ByVal pb As PictureBox)
    ' En lugar de delegar en la interfaz de IShape
    '   del Polygon interno,
    '   Rectangle aprovecha el hecho de
    '   que hay un comando gráfico
    '   que dibuja un cuadro en una sola operación
    '   en vez de en cuatro (los gráficos
    '   son presumiblemente la clase de operación
    '   que consume más recursos).
    Dim sngX1 As Single, sngY1 As Single
    Dim sngX2 As Single, sngY2 As Single
    Call mpyg.GetPoint(0, sngX1, sngY1)
    Call mpyg.GetPoint(1, sngX2, sngY2)
    If mblnFilled Then
        pb.Line (sngX1, sngY1)-(sngX2, sngY2), mrgbColor, BF
    Else
        pb.Line (sngX1, sngY1)-(sngX2, sngY2), mrgbColor, B
    End If
End Sub

' El método IShape.TimeTest se utiliza para mostrar
'   el menor coste de recursos
'   de la llamada a un método de una
'   interfaz implementada por varias clases
'   en oposición a la llamada a un
'   método similar de las interfaces
'   predeterminadas de las clases.
'
Private Sub IShape_TimeTest()
End Sub

' --------------------------------------------------------
' Este es el principio de la implementación
'   de la interfaz Polygon.
' Polygon.Color - Como no se utiliza un
'   Polygon interno para
'   dibujar el objeto Rectangle, la propiedad
'   Color queda completamente ignorada.
'   Esto no es necesario; Rectangle podría
'   delegar el almacenamiento de la propiedad
'   Color en el Polygon interno (como
'   hace la clase Triangle), ahorrando
'   código de implementación y almacenamiento.
'   El color podría recuperarse
'   desde el Polygon interno
'   al terminar de dibujar.
Private Property Get Polygon_Color() As Long
    Polygon_Color = mrgbColor
End Property
'
Private Property Let Polygon_Color(ByVal RHS As Long)
    If 0 <> (RHS And &HFF000000) Then
        Err.Raise vbObjectError + 2080, , _
            "Valor no válido de color para Polygon."
        Exit Property
    End If
    mrgbColor = RHS
End Property

' Polygon.TimeTest - Rectangle tiene tres
'   métodos TimeTest, uno
'   en la interfaz de IShape (utilizado para mostrar
'   el polimorfismo y la asociación previa), otro en su
'   propia interfaz (utilizado para mostrar la asociación
'   tardía), y ésta. Esta es un producto del
'   efecto secundario del hecho de que Rectangle
'   implementa la interfaz de Polygon; no se
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
    '   Polygon interno, puesto que
    '   Rectangle siempre está especificado
    '   por dos puntos solamente.
    Polygon_GetPointCount = 2
End Property

' Polygon.SetPoints - Al implementar el
'   método SetPoints
'   de la interfaz Polygon, Rectangle
'   ejecuta su propio código para verificar que
'   la matriz de entrada sólo contiene dos
'   puntos (cuatro Singles), y después
'   delega en el objeto Polygon
'   interno.
Private Sub Polygon_SetPoints(asngPoints() As Single)
    Dim blnBadArray As Boolean
    On Error Resume Next
    ' Asegura que la matriz de entrada no contiene
    '   más de cuatro puntos. (El
    '   método SetPoint de Polygon verificará
    '   que la matriz tiene
    '   base cero.)
    If UBound(asngPoints) <> 3 Then blnBadArray = True
    ' Si al llamar a Ubound se produce un error,
    '   rechaza la matriz.
    If Err.Number <> 0 Then blnBadArray = True
    If blnBadArray Then
        Err.Raise vbObjectError + 2083, , _
            "Un rectángulo se especifica con una matriz de cuatro números (izquierda, arriba, derecha, abajo) en una matriz basada en cero."
        Exit Sub
    End If
    ' Delega en el Polygon interno, que
    '   completa la validación de la matriz y
    '   la almacena.
    Call mpyg.SetPoints(asngPoints)
End Sub

' ----------------------------------------------------------------
' Este es el principio de la propia interfaz (predeterminada)
'   del objeto Rectangle.

' El método TimeTest no tiene
'   argumentos, y
'   vuelve inmediatamente.  Se utiliza para
'   ilustrar la asociación tardía, en oposición
'   a la asociación previa proporcionada por las
'   llamadas a TimeTest en la interfaz de
'   IShape.
Public Sub TimeTest()
End Sub

' La propiedad Filled determina si el
'   rectángulo está lleno al
'   dibujarlo.
Public Property Get Filled() As Boolean
    Filled = mblnFilled
End Property
'
Public Property Let Filled(ByVal NewValue As Boolean)
    mblnFilled = NewValue
End Property

' ----------------------------------------------------------------
' Este es el principio de los procedimientos
'   privados de la clase (procedimientos de ayuda
'   y procedimientos de evento).

Private Sub Class_Initialize()
    Dim asngPoints(0 To 3) As Single
    ' Código de depuración.
    mlngDebugID = DebugInit(Me)
    '
    ' Crea el objeto Polygon interno, y
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

