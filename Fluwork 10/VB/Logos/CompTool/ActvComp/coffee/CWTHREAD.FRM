VERSION 5.00
Begin VB.Form frmThread 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demostración de múltiples subprocesos"
   ClientHeight    =   4950
   ClientLeft      =   630
   ClientTop       =   690
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7095
   Begin VB.Timer tmrShort 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   3840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar tareas largas"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtN 
      Height          =   375
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "4"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdNLong 
      Caption         =   "...con &N tareas largas"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdShortLong 
      Caption         =   "... c&on tareas largas"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton cmdShortOnly 
      Caption         =   "&Tareas cortas, serializadas..."
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdIDs 
      Caption         =   "&Mostrar Id. de los subprocesos"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.ListBox lstResults 
      Height          =   4575
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdXThread 
      Caption         =   "&Sobrecarga entre subprocesos"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Sus pruebas son..."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   2895
   End
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' La demostración de múltiples subprocesos:
'   - Muestra los id del subroceso y el número de
'       objetos de un subroceso (esto será más
'       interesante si compila MTCoffee.exe
'       con un tamaño de pool de subrocesos de 3 o 4).
'   - Compara la complejidad de procesamiento para
'       las llamadas en un solo subroceso y entre
'       varios subrocesos.
'   - Tiempo por iteración para tareas cortas en
'       serie (vea el objeto Coffee definido en
'       MTCoffee.cls).
'   - Tiempo por iteración para una tarea larga, y
'       para tareas cortas en serie ejecutadas al
'       mismo tiempo.
'   - Tiempo por iteración para varias tareas largas,
'       con tareas cortas en serie ejecutándose al
'       mismo tiempo.
'
' Puede agregar sus propias pruebas a este esquema.
'   Puede encontrar interesante comparar el comportamiento
'   de las tareas que bloquean (como consultas a bases de
'   datos de sistemas remotos, o transferencias de archivos
'   grandes). En un equipo con un solo procesador, tales
'   tareas se realizarán mucho mejor que las de cálculo
'   intenso de los ejemplos anteriores. Los subrocesos que
'   realizan tareas de cálculo intenso compiten entre sí
'   por el único procesador de la máquina, por lo que su
'   rendimiento se resiente al aumentar el número de
'   subrocesos activos.
'
' Si desea más información, vea "Escalabilidad y subrocesos
'   múltiples", en "Generación de componentes de código", en los
'   Libros en pantalla.
 
' Estas constantes controlan el tamaño relativo de una tarea
'   larga y de una tarea corta. Puede ser necesario ajustarlas
'   a la velocidad de su procesador.
Const SHORTTASKSIZE = 50000
Const LONGTASKSIZE = 2000000

' Matriz de objetos Coffee.
Private macfe(1 To 20) As Coffee

' Colección de objetos CoffeeTracker.
Public CoffeeTrackers As New Collection

' Indicador de cancelar.
Public CancelAll As Boolean

' Número de tareas cortas a ejecutar.
Private mintHowManyShort As Integer

Private Sub cmdCancel_Click()
    ' Cuando el usuario hace clic en Cancelar,
    '   establece un indicador que CoffeeTracker
    '   puede usar para cancelar todas las tareas
    '   largas la siguiente vez que éstas produzcan
    '   un evento Progress.
    CancelAll = True
End Sub

' Ejecuta algunas tareas cortas en serie, para tener
'   idea de su velocidad cuando no compiten por el
'   procesador.
Private Sub cmdShortOnly_Click()
    lstResults.Clear
    CancelAll = False
    mintHowManyShort = 10
    tmrShort.Interval = 250
    tmrShort.Enabled = True
End Sub

' Ejecuta una tarea larga, y luego tareas
'   cortas en serie mientras se ejecuta la
'   primera.
Private Sub cmdShortLong_Click()
    Dim cft As CoffeeTracker
    
    lstResults.Clear
    CancelAll = False
    Set cft = NewTracker(macfe(1).ThreadID, LONGTASKSIZE)
    Set cft.Coffee = macfe(1)
    Call macfe(1).StartLongTask(LONGTASKSIZE)
    '
    ' Organiza algunas tareas cortas para que se
    '   ejecuten (una cada cuarto de segundo)
    '   mientras se ejecuta la tarea.
    mintHowManyShort = 10
    tmrShort.Enabled = True
    tmrShort.Interval = 250
End Sub

' Inicia N tareas largas (1 - 9), y luego ejecuta una
'   serie de tareas cortas.
Private Sub cmdNLong_Click()
    Dim intCt As Integer
    Dim cft As CoffeeTracker
    
    lstResults.Clear
    CancelAll = False
    For intCt = 1 To CLng("0" & txtN)
        Set cft = NewTracker(macfe(intCt).ThreadID, LONGTASKSIZE)
        Set cft.Coffee = macfe(intCt)
        Call macfe(intCt).StartLongTask(LONGTASKSIZE)
    Next
    '
    ' Organiza algunas tareas cortas para que se
    '   ejecuten mientras lo hacen las largas.
    mintHowManyShort = 10 + CLng("0" & txtN)
    tmrShort.Enabled = True
    tmrShort.Interval = 250
End Sub

' Compara el tráfico de las llamadas a un objeto
'   en el mismo subroceso con el de las llamadas a
'   un objeto en otro subroceso.
'
Private Sub cmdXThread_Click()
    Dim cfeSame As Coffee
    Dim cfeNew As Coffee
    
    cmdXThread.Caption = "Trabajando..."
    cmdXThread.Enabled = False
    
    ' Crea un objeto Coffee en otro subroceso.
    Set cfeNew = macfe(10).GetCoffeeOnNewThread
    ' En caso de pool de subrocesos, excepto en
    '   caso de generado de un solo subroceso en
    '   el pool, debe obtenerse un subroceso distinto
    '   la segunda vez.
    If cfeNew.ThreadID = macfe(10).ThreadID Then
        Set cfeNew = macfe(10).GetCoffeeOnNewThread
        If cfeNew.ThreadID = macfe(10).ThreadID Then
            MsgBox "Imposible establecer la comparación entre llamadas en el mismo subproceso y entre subprocesos; imposible obtener un objeto de otro subproceso."
            Exit Sub
        End If
    End If
    '
    ' Crea un objeto Coffee en el mismo subroceso.
    Set cfeSame = macfe(10).GetCoffeeOnSameThread
    
    ' Usa los objetos Coffee recién creados para
    '   realizar la prueba.
    MsgBox "Mismo subproceso:  " & macfe(10).CallAnotherCoffee(cfeSame) & " seg/llamada" & vbCrLf _
        & "Entre subprocesos:  " & macfe(10).CallAnotherCoffee(cfeNew) & " seg/llamada"
    
    cmdXThread.Caption = "Sobrecarga entre subprocesos"
    cmdXThread.Enabled = True
    '
    ' Los objetos Coffee creados para esta prueba
    '   se terminan cuando cfeNew y cfeSame
    '   quedan fuera de ámbito al final de este
    '   procedimiento.
End Sub

' Enumera los ID de subroceso de los objetos Coffee
'   creados al cargarse el formulario.
'
Private Sub cmdIDs_Click()
    Dim intCt As Integer
    lstResults.Clear
    For intCt = 1 To 10
        lstResults.AddItem macfe(intCt).ThreadID _
            & "  (en subproceso " & macfe(intCt).NumberOnThread & ")"
    Next
End Sub

' Crea una serie de objetos Coffe, cada uno en su propio subroceso.
'   (Si vuelve a compilar MTCoffee con un pool de menos de
'   10 subrocesos, algunos de los objetos compartirán el subroceso y
'   el estado global).
'
Private Sub Form_Load()
    Dim intCt As Integer
    For intCt = 1 To 10
        Set macfe(intCt) = New Coffee
    Next
    Form1.cmdMT.Enabled = True
    Form1.cmdMT.MousePointer = vbDefault
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 8
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

' Proporciona claves únicas para los CoffeeTrackers.
'
Private Function NewKey() As String
    Static lngLastKey As Long
    lngLastKey = lngLastKey + 1
    NewKey = "K" & lngLastKey
End Function

' Agrega un nuevo CoffeeTracker. Con rigor,
'   este debería ser un método de una clase
'   colección CoffeeTrackers.
'
Private Function NewTracker(ByVal ThreadID As Long, _
        ByVal Size As Long) As CoffeeTracker
        
    Dim cft As New CoffeeTracker
    '
    ' Almacena el ID de subroceso del objeto Coffee
    '   del que se va a hacer el seguimiento.
    cft.ThreadID = ThreadID
    '
    ' Establece el tamaño de la tarea asignada al
    '   objeto Coffee del que se va a hacer el
    '   seguimiento.
    cft.Size = Size
    '
    ' Da al seguidor una clave única para la
    '   colección.
    cft.ID = NewKey
    '
    ' Incluye el nuevo seguidor en una colección.
    CoffeeTrackers.Add cft, cft.ID
    '
    ' Devuelve una referencia al nuevo seguidor.
    Set NewTracker = cft
End Function

' El cronómetro se usa para iniciar una serie
'   de tareas cortas a intervalos regulares,
'   usando un objeto Coffee (es decir un solo
'   subroceso). Antes de empezar una nueva tarea,
'   comprueba si ha terminado la anterior.
'
' Si ejecuta MTCoffee en el entorno de
'   desarrollo, marque como comentario el código
'   que comprueba si ha terminado la tarea anterior.
'
Private Sub tmrShort_Timer()
    Static intCt As Integer
    Static strWaitingFor As String
    Dim cft As CoffeeTracker
    
    ' Si strWaitingFor contiene una clave, el
    '   objeto Coffee está (o estaba) realizando
    '   una tarea, y la variable de cadena estática
    '   contiene la clave del CoffeeTracker
    '   que lo está observando.
    If strWaitingFor <> "" Then
        On Error Resume Next
        '
        ' Si el CoffeeTracker que esperamos ha
        '   quedado fuera de la colección, se
        '   producirá un error, lo que implica
        '   que es el momento de iniciar otro.
        Set cft = CoffeeTrackers(strWaitingFor)
        If Err.Number = 0 Then Exit Sub
    Else
        ' Si la variable de cadena estática
        '   está vacía, la serie anterior de
        '   tareas cortas ha terminado. Comienza
        '   una nueva serie, por lo que se
        '   restablece el contador (intCt).
        intCt = 0
    End If
    
    If Not CancelAll Then
        intCt = intCt + 1
        '
        ' Crea un CoffeeTracker para esperar
        '   a que termine la petición.
        Set cft = NewTracker(macfe(10).ThreadID, SHORTTASKSIZE)
        '
        ' Da al CoffeeTracker el objeto
        '   Coffee a observar.
        Set cft.Coffee = macfe(10)
        '
        ' Empieza la tarea.
        Call macfe(10).StartLongTask(SHORTTASKSIZE)
        '
        ' Se prepara para esperar
        '   al CoffeeTracker.
        strWaitingFor = cft.ID
    End If
    '
    ' Comprueba si se ha completado la serie
    '   de tareas cortas, o si se ha detenido
    '   con el botón Cancelar:
    If (intCt >= mintHowManyShort) Or CancelAll Then
        intCt = 0
        tmrShort.Enabled = False
        strWaitingFor = ""
    End If
End Sub
