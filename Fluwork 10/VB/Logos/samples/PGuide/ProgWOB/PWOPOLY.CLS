VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Polygon"
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

' Polygon es una clase MUY rudimentaria de
'   objeto polígono.  Sólo
'   permite crear polígonos (con el
'   método SetPoints); no hay manera de
'   modificar los puntos una vez establecidos.
'   (Por supuesto, puede volver a llamar a SetPoints
'   para volver a definir el polígono).

Private Type POLYPOINT  ' prefijo pyp
    X As Single
    Y As Single
    Angle As Single
End Type

' El Polygon predeterminado tiene un punto, en (0,0).
Private mapyp() As POLYPOINT

' Almacenamiento para la propiedad Normalized de sólo lectura.
Private mblnNormalized As Boolean

' Almacenamiento para la propiedad Color.
Private mrgbColor As Long

' IShape es la interfaz que se utiliza para
'   presentar el polígono. También
'   tiene un método TimeTest que se utiliza para
'   comparar el coste de la asociación de tipos
'   previa y tardía.
Implements IShape

' -------------------------------------
' Esto marca el principio de la
'   implementación de la
'   interfaz IShape.

' Se llama a IShape.DrawToPictureBox
'   para dibujar una forma,
'   de forma que cada clase de forma tiene que
'   aportar su propia implementación.
'
Private Sub IShape_DrawToPictureBox(ByVal pb As PictureBox)
    Dim sngXLast As Single
    Dim sngYLast As Single
    Dim sngX As Single
    Dim sngY As Single
    Dim intMax As Integer
    Dim intCt As Integer
    
    intMax = UBound(mapyp)
    sngX = mapyp(intMax).X
    sngY = mapyp(intMax).Y
    If intMax = 0 Then
        pb.PSet (sngX, sngY), mrgbColor
    Else
        For intCt = 0 To intMax
            sngXLast = sngX
            sngYLast = sngY
            sngX = mapyp(intCt).X
            sngY = mapyp(intCt).Y
            pb.Line (sngXLast, sngYLast)-(sngX, sngY), mrgbColor
        Next
    End If
End Sub

' El método IShape.TimeTest se utiliza para mostrar
'   la reducción del coste
'   de la llamada a un método que es llamado
'   en una interfaz implementada por
'   varias clases, en oposición a llamar
'   a un método similar en las
'   interfaces predeterminadas de las clases.
'
Private Sub IShape_TimeTest()
End Sub

' --------------------------------------
' Este es el principio de la interfaz
'   predeterminada de la clase Polygon (propiedades Public
'   y métodos). Esta es la interfaz
'   de Polygon implementada por Triangle
'   y Rectangle.

' Propiedad Color.
' -----
'
Public Property Get Color() As Long
    Color = mrgbColor
End Property
'
Public Property Let Color(ByVal rgb As Long)
    If 0 <> (rgb And &HFF000000) Then
        Err.Raise vbObjectError + 2080, , _
            "Valor no válido de color para Polygon."
        Exit Property
    End If
    mrgbColor = rgb
End Property

' El método TimeTest no tiene argumentos,
'   no hace nada, y
'   vuelve inmediatamente. Es necesario para
'   ilustrar el coste de las llamadas de
'   asociación tardía, en oposición a la
'   asociación previa proporcionada por
'   TimeTest en la interfaz de IShape.
'
' Podría pensar que podríamos hacer TimeTest
'   una propiedad Friend, como DebugID, para guardar
'   Triangle y Rectangle que
'   implementan la interfaz de Polygon por
'   tener que implementar Polygon_TimeTest.
'   (Las propiedades Friend y los métodos NO forman
'   parte de la interfaz de una clase).
'   La razón por la que no lo hacemos es porque
'   TimeTest tiene que ser llamado con asociación tardía
'   en la demostración mientras que las propiedades Friend
'   y los métodos siempre son llamados con
'   asociación previa.
Public Sub TimeTest()
End Sub

' GetPoint establece dos Singles ByRef a los
'   valores X e Y del punto
'   requerido. (Si la interfaz de Polygon
'   no fuera implementada por
'   Triangle y Rectangle, GetPoint
'   podría declararse Friend, y podría
'   devolver un POLYPOINT que tendría que
'   ser declarado Public en un
'   módulo estándar en ese caso;
'   sin embargo, los miembros Friend no forman parte
'   de la interfaz de una clase, así que hacer
'   GetPoint un miembro Friend impediría que
'   Triangle y Rectangle implementasen
'   una asociación previa en
'   Polygon_GetPoint).
'
Public Sub GetPoint(ByVal intPoint As Integer, _
        ByRef X As Single, ByRef Y As Single)
    X = mapyp(intPoint).X
    Y = mapyp(intPoint).Y
End Sub

' GetPointCount devuelve el número de
' puntos del objeto Polygon.
'
Public Property Get GetPointCount() As Integer
    GetPointCount = UBound(mapyp) + 1
End Property

' SetPoints acepta una matriz de Singles
'   con base cero, los elementos con números
'   pares (0, 2, etc.) son los valores X,
'   y los elementos impares
'   son los valores Y de los puntos.
'
Public Sub SetPoints(asngPoints() As Single)
    Dim blnBadArray As Boolean
    Dim intMax As Integer
    Dim intPoint As Integer
    
    On Error Resume Next
    If LBound(asngPoints) <> 0 Then blnBadArray = True
    intMax = UBound(asngPoints)
    ' El límite superior de una matriz con base cero
    '   tiene que ser un número impar,
    '   ésto lo valida.
    If (intMax / 2#) = (intMax \ 2) Then blnBadArray = True
    ' Si en la función Ubound se produce un error,
    ' declare no válida la matriz.
    If Err.Number <> 0 Then blnBadArray = True
    If blnBadArray Then
        Err.Raise vbObjectError + 2081, , _
            "SetPoints debe recibir una matriz basada en cero de una dimensión con un número par de elementos, las entradas impares son valores X y la entradas pares son valores Y."
        Exit Sub
    End If
    ' Convierte el índice máximo de la matriz
    '   de entrada al índice máximo de la matriz
    '   interna del objeto Polygon.
    intMax = intMax \ 2
    ReDim mapyp(0 To intMax)
    ' Lee los valores de los puntos.
    For intPoint = 0 To intMax
        mapyp(intPoint).X = asngPoints(intPoint * 2)
        mapyp(intPoint).Y = asngPoints(intPoint * 2 + 1)
    Next
End Sub

' ------------------------------------------------------
' Este es el principio de los procedimientos privados
'   de Polygon (procedimientos de ayuda
'   y procedimientos de evento).

Private Sub Class_Initialize()
    ' Código de depuración.
    mlngDebugID = DebugInit(Me)
    '
    ' El Polygon predeterminado tiene un punto.
    ReDim mapyp(0 To 0)
End Sub

Private Sub Class_Terminate()
    DebugTerm Me
End Sub

' -------- Implementación de IDebug --------
'
' IDebug.DebugID ofrece la manera de separar
'   los objetos. Es requerido por los
'   procedimientos de depuración
'   DebugInit, DebugTerm y DebugShow,
'   declarados en modFriend.
'
Private Property Get IDebug_DebugID() As Long
    IDebug_DebugID = mlngDebugID
End Property

