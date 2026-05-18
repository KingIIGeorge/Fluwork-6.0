VERSION 5.00
Begin VB.Form frmImplements 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polimorfismo y la palabra clave Implements"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "PWOImple.frx":0000
   ScaleHeight     =   3510
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDisplay 
      Interval        =   220
      Left            =   1800
      Top             =   2760
   End
   Begin VB.PictureBox picShapes 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      Picture         =   "PWOImple.frx":0446
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdLate 
      Caption         =   "E&nlace en tiempo de ejecución"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3000
      Width           =   2595
   End
   Begin VB.CommandButton cmdEarly 
      Caption         =   "&Enlace en tiempo de compilación"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2520
      Width           =   2595
   End
   Begin VB.Label lblLateResult 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblEarlyResult 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Sobrecarga por llamada a métodos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"PWOImple.frx":088C
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6075
   End
End
Attribute VB_Name = "frmImplements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NUMOBJECTS = 100
Const NUMREPEATSEARLY As Long = 10000
Const NUMREPEATSLATE As Long = 500

' Por razones de la demostración, se mantienen tres matrices
'   de referencias de objetos.  Cada objeto creado por la
'   demostración tiene una entrada en las tres
'   matrices:
' Matriz de interfaces IShape;
Private maishEarly(1 To NUMOBJECTS) As IShape
' Matriz de interfaces Polygon;
Private mapyg(1 To NUMOBJECTS) As Polygon
' Matriz de interfaces predeterminadas (Polygon,
'   Triangle o Rectangle, dependiendo del
'   objeto).
Private maobjLate(1 To NUMOBJECTS) As Object

' Prueba de tiempo de las llamadas de asociación previa
'   de los objetos NUMOBJECTS,
'   utilizando la interfaz IShape. (El método
'   TimeTest es sólo para consumir,
'   no tiene argumentos, y vuelve inmediatamente.)
'   Esto se repite
'   NUMREPEATSEARLY veces.
'
Private Sub cmdEarly_Click()
    Dim lngCt As Long
    Dim intCt As Integer
    Dim timeMark As Long
    
    ' Desactiva la presentación de formas durante
    '   la prueba.
    tmrDisplay.Enabled = False
    cmdEarly.Caption = "Trabajando..."
    timeMark = timeGetTime
    For lngCt = 1 To NUMREPEATSEARLY
        For intCt = 1 To NUMOBJECTS
            ' Hace las llamadas a TimeTest
            '   a través de la interfaz IShape,
            '   implementada por las tres clases
            '   (Polygon, Rectangle y
            '   Triangle).
            maishEarly(intCt).TimeTest
        Next
    Next
    timeMark = timeGetTime - timeMark
    lblEarlyResult = ShowElapsed(timeMark, _
        NUMOBJECTS * NUMREPEATSEARLY, _
        "Sobrecarga de llamadas por enlace en tiempo de compilación")
    cmdEarly.Caption = "&Enlace en tiempo de compilación"
    '
    ' Vuelve a presentar las formas.
    tmrDisplay.Enabled = True
End Sub

Private Sub cmdLate_Click()
    Dim lngCt As Long
    Dim intCt As Integer
    Dim timeMark As Long
    
    ' Desactiva la presentación de las formas durante
    '   la prueba.
    tmrDisplay.Enabled = False
    cmdLate.Caption = "Trabajando..."
    timeMark = timeGetTime
    For lngCt = 1 To NUMREPEATSLATE
        For intCt = 1 To NUMOBJECTS
            ' Hace llamadas a TimeTest
            '   con asociación tardía, a través de las
            '   interfaces predeterminadas de los objetos
            '   (Polygon, Rectangle y
            '   Triangle).
            maobjLate(intCt).TimeTest
        Next
    Next
    timeMark = timeGetTime - timeMark
    lblLateResult = ShowElapsed(timeMark, _
        NUMOBJECTS * NUMREPEATSLATE, _
        "Sobrecarga de llamadas por enlace en tiempo de ejecución")
    cmdLate.Caption = "E&nlace en tiempo de ejecución"
    '
    ' Vuelve a presentar las formas.
    tmrDisplay.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Libera los recursos asociados con el
    '   formulario, borrando la variable
    '   global oculta.
    Set frmImplements = Nothing
End Sub

' Si se hace clic en el cuadro de imagen
'   de las formas, presenta todas las formas a
'   la vez (utilizando asociación previa).
Private Sub picShapes_Click()
    Dim intCt As Integer
    For intCt = 1 To NUMOBJECTS
        Call maishEarly(intCt).DrawToPictureBox(picShapes)
    Next
End Sub

Private Sub Form_Load()
    Dim intCt As Integer
    Dim asngPoints() As Single
    Set Picture = New StdPicture
    Call Randomize(Timer)
    For intCt = 1 To NUMOBJECTS
        ' Crea aleatoriamente entre 1 y 3 Polígonos,
        '    entre 1 y 3Triangles y entre 1 y 3 Rectángulos.
        '   Almacena la referencia de cada uno de
        '   estos objetos en la matriz de asociación
        '   tardía.
        Select Case Int(Rnd * 3)
            Case 0
                Set maobjLate(intCt) = New Polygon
            Case 1
                Set maobjLate(intCt) = New Triangle
            Case 2
                Set maobjLate(intCt) = New Rectangle
        End Select
        ' Guarda una referencia a la interfaz del objeto
        '   IShape, para demostrar la asociación previa
        '   mediante polimorfismo.
        '   Las tres clases
        '   implementan IShape, así que Visual
        '   Basic puede consultar la
        '   interfaz IShape y hacer la
        '   asignación.
        Set maishEarly(intCt) = maobjLate(intCt)
        ' También guarda una referencia a la interfaz
        '   del objeto Polygon.
        Set mapyg(intCt) = maobjLate(intCt)
        ' Si el objeto era Polygon (en lugar de
        '   Triangle o Rectangle,
        '   que simplemente implementan
        '   la interfaz de Polygon), sólo tendrá
        '   un punto.  Asígnele un número
        '   aleatorio de puntos (entre
        '   4 y 24).
        If mapyg(intCt).GetPointCount = 1 Then
            ReDim asngPoints(0 To (Int(21 * Rnd) + 4) * 2 - 1)
            Call mapyg(intCt).SetPoints(asngPoints)
        End If
        ' Asigna un color aleatorio al objeto.
        mapyg(intCt).Color = Int(Rnd * &HFFFFFF)
    Next
    Debug.Print "Si mira atrás en los números de depuración"
    Debug.Print "de los objetos shape, se dará cuenta de que se "
    Debug.Print "crearon muchos más objetos que " & NUMOBJECTS & "."
    Debug.Print "Esto se debe a que cada Triangle y cada"
    Debug.Print "Rectangle se crearon en su propio objeto Polygon."
    '
    ' Asigna valores aleatorios a cada punto
    '   de cada objeto.
    Call NewShapes
End Sub

Private Sub tmrDisplay_Timer()
    ' Recorre repetidamente los objetos
    '   Shape, presentándolos en un
    '   cuadro de imagen.
    Static intCt As Integer
    picShapes.Cls
    intCt = intCt + 1
    If intCt > NUMOBJECTS Then intCt = 1
    Call maishEarly(intCt).DrawToPictureBox(picShapes)
End Sub

' El procedimiento ShowElapsed presenta
' -----------       el resultado de una prueba de tiempo
'   de asociación previa o tardía.
'
Private Function ShowElapsed(ByVal Milliseconds As Long, _
        ByVal Iterations As Long, _
        ByVal Caption As String) As String
    Dim strMSec As String
    strMSec = Format$(Milliseconds / Iterations, "0.0000")
    
    MsgBox Format$(Iterations, "#,###,##0") _
        & " iteraciones en " _
        & Format$(Milliseconds / 1000#, "##,##0.00") _
        & " segundos" & vbCrLf _
        & strMSec _
        & " milisegundos por llamada", , Caption
        
    ShowElapsed = strMSec & " mseg/llamada"
End Function

' NewShapes cambia la forma de cada uno de
' ---------     los objetos, llamando a
'   MakeRandomPoints para generar un conjunto de
'   puntos aleatorios.  No modifica el
'   el número de puntos de un objeto Polygon.
'
Private Sub NewShapes()
    Dim intCt As Integer
    Dim intPt As Integer
    Dim pyg As Polygon
    Dim asngPoints() As Single
    Dim intNumPts As Integer
    
    For intCt = 1 To NUMOBJECTS
        intNumPts = mapyg(intCt).GetPointCount
        Call MakeRandomPoints(intNumPts, asngPoints)
        Call mapyg(intCt).SetPoints(asngPoints)
    Next
End Sub

' MakeRandomPoints crea un conjunto de puntos
' ----------------      aleatorios para un
'   objeto Polygon, colocándolos en la matriz
'   unidimensional con base cero que
'   requiere el método SetPoints.
'
Private Sub MakeRandomPoints( _
        ByVal intNumPts As Integer, _
        asngPoints() As Single)
        
    Dim intPt As Integer
    
    ReDim asngPoints(0 To intNumPts * 2 - 1)
    For intPt = 0 To intNumPts * 2 - 1 Step 2
        asngPoints(intPt) = Rnd * picShapes.ScaleWidth
        asngPoints(intPt + 1) = Rnd * picShapes.ScaleHeight
    Next
End Sub

