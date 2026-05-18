VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coffee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

' El objeto Coffee representa un estilo distinto
'   de notificaciones asíncronas con respecto a
'   las que realizan los CofeeMonitors. En lugar
'   de notificaciones periódicas, Coffee proporciona
'   informes sobre una tarea larga, y un evento de
'   terminación.
'
' El mecanismo empleado para estas notificaciones
'   es provocar eventos. También pueden usarse
'   métodos de llamada de retorno, y de hecho
'   ello supone ciertas ventajas. Las llamadas de
'   retorno permiten a un componente tratar de
'   modo inteligente los errores del cliente,
'   mientras que los eventos no devuelven errores
'   de cliente. Este tema se trata en "Cuándo
'   utilizar eventos o devoluciones de llamada para
'   las notificaciones", en los Libros en pantalla.

' Número de iteraciones a realizar en la tarea
'   de relleno.
Private mlngIterations As Long

' XTimer se usa para iniciar la tarea
'   larga de forma asíncrona.
Private WithEvents mwXTimer As XTimer
Attribute mwXTimer.VB_VarHelpID = -1

Event Progress(ByVal PercentDone As Single, _
        ByRef Cancel As Boolean)
Event Complete(ByVal Canceled As Boolean)
        
' ThreadID devuelve el ID de subproceso de sistema
' --------      en el que se ha credo el objeto.
'
Public Property Get ThreadID() As Long
    ThreadID = App.ThreadID
End Property

' NumberOnThread devuelve el número de objetos
' --------------    Coffee en ejecución en este
'   subproceso. Este número es el valor de la variable
'   de datos global glngGlobalData, la cual
'   incrementan los objetos Coffee en su evento
'   Initialize y decrementan en su evento
'   Terminate.
'
' Si se ha compilado MTCoffee con Subproceso por
'   objeto, la única forma de que varios objetos
'   compartan un subproceso (y la instancia de datos
'   globales asociada a él) se da cuando se crea
'   otro objeto Coffee en el subproceso llamando a
'   GetCoffeeOnSameThread.
'
' Si MTCoffee se ha compilado con un pool de
'   subprocesos y la cuenta de objetos activos
'   excede el número de subprocesos del pool, los
'   objetos Coffee compartirán subprocesos.
'
Public Property Get NumberOnThread() As Long
    NumberOnThread = glngGlobalData
End Property

' StartLongTask prepara lo necesario para la
' -------------     tarea de relleno larga. La
'   tarea la inicia en realidad un XTimer sólo
'   de código que StartLongTask pone en
'   ejecución.
'
Public Sub StartLongTask(ByVal Iterations As Long)
    ' Esta es una forma abreviada para probar
    '   el tiempo que suponen las llamadas.
    '   Vea CallAnotherCoffee.
    If Iterations = 0 Then Exit Sub
    '
    ' Almacena el tamaño de la tarea de relleno.
    mlngIterations = Iterations
    '
    ' Da al cronómetro un intervalo corto,
    '   y lo pone a funcionar inmediatamente
    '   antes de regresar.
    mwXTimer.Interval = 55
    mwXTimer.Enabled = True
End Sub

' GetCoffeeOnSameThread crea un nuevo objeto
' ---------------------     Coffee en el mismo
'   subproceso, simulando los efectos de un pool de
'   subprocesos. Esto sólo puede hacerse internamente,
'   como se explica en "Funcionamiento de la
'   creación de objetos en Visual Basic", en los
'   Libros en pantalla.
'
Public Function GetCoffeeOnSameThread() As Coffee
    ' Todos los objetos creados con New estarán
    '   en el subproceso de su creador, incluso en
    '   el caso de un nuevo objeto Coffee.
    Set GetCoffeeOnSameThread = New Coffee
End Function

' GetCoffeeOnNewThread crea un nuevo objeto
' --------------------     Coffee en un nuevo
'   subproceso, llamando a CreateObject para crear
'   el nuevo objeto Coffee. La diferencia con
'   la creación interna de GetCoffeeOnSameThread
'   se explica en "Funcionamiento de la creación
'   de objetos en componentes de Visual Basic",
'   en los Libros en pantalla.
'
' Observe que esta técnica podría emplearse para
'   crear objetos en subprocesos diferentes que
'   podrían comunicarse entre sí, sin que el
'   cliente tuviese que pasar a un objeto una
'   referencia al otro (como hace CoffeeWatch).
'   Si experimenta con ello, recuerde que el
'   coste de transferir llamadas entre subprocesos
'   es casi tan grande como el de hacerlo entre
'   procesos.
'
Public Function GetCoffeeOnNewThread() As Coffee
    ' Lo crea como lo haría un cliente externo.
    Set GetCoffeeOnNewThread = CreateObject("MTCoffee.Coffee")
End Function

' CallAnotherCoffee da una medida aproximada
' -----------------     del tiempo de la llamada
'   entre procesos. Pásele un objeto Coffee de
'   otro subproceso, o del mismo, y compare el
'   resultado. El método hace llamadas de relleno
'   a StartLongTask, por lo que en esencia sólo
'   está midiendo el tiempo de la llamada.
'
Public Function CallAnotherCoffee(ByVal cfe As Coffee) As Double
    Const TRIES = 10000
    Dim timeStart As Long
    Dim timeEnd As Long
    Dim lngTries As Long
    
    timeStart = timeGetTime
    For lngTries = 1 To TRIES
        cfe.StartLongTask 0
    Next
    timeEnd = timeGetTime
    '
    ' Devuelve los segundos (ss.mmm) por
    '   llamada (el resultado será incorrecto
    '   si ejecuta CallAnotherCoffee inmediatamente
    '   después de que el cronómetro del sistema
    '   haya sobrepasado el cero).
    CallAnotherCoffee = ((CDbl(timeEnd) - timeStart) / 1000#) / TRIES
End Function

Private Sub Class_Initialize()
    ' Incrementa la cuenta global (es decir,
    '   la de este subproceso) de objetos Coffee.
    glngGlobalData = glngGlobalData + 1
    '
    ' Crea un objeto cronómetro.
    Set mwXTimer = New XTimer
End Sub

Private Sub Class_Terminate()
    ' Decrementa la cuenta global (es decir,
    '   la de este subproceso) de objetos Coffee.
    glngGlobalData = glngGlobalData - 1
    '
    ' Libera el objeto cronómetro.
    Set mwXTimer = Nothing
End Sub

Private Sub mwXTimer_Tick()
    ' Lo primero, desactivar el cronómetro.
    mwXTimer.Enabled = False
    Call LongTask
End Sub

' Tarea de relleno.
'
Private Sub LongTask()
    Dim dblDummy As Double
    Dim lngCt As Long
    Dim sngNextMark As Single
    Dim blnCancel As Boolean
    
    ' Para transacciones pequeñas, no se
    '   preocupa de hacer la llamada de
    '   retorno durante la ejecución.
    If mlngIterations < 100000 Then
        sngNextMark = 1!
    Else
        sngNextMark = 0.1!
    End If
        
    ' Esto es sólo para perder el tiempo.
    For lngCt = 1 To mlngIterations
        ' Si esto fuese una aplicación real,
        '   aquí se realizaría una unidad de
        '   trabajo. Puede encontrar interesante
        '   reemplazar esta actividad de ocupación
        '   intensiva del procesador por otra que
        '   haga esperar mucho al sistema, como
        '   las llamadas a una base de datos de
        '   otra máquina, o la lectura de un
        '   archivo muy grande. El rendimiento en
        '   una estación de trabajo con un solo
        '   procesador es mucho mayor cuando la
        '   mayoría de los subprocesos están
        '   bloqueados, esperando una entrada de
        '   archivo o el resultado de una llamada
        '   a una base de datos.
        '
        dblDummy = 3033.14159 * 2081.14159 * 1138.14159
        '
        If CDbl(lngCt) / mlngIterations > sngNextMark Then
            RaiseEvent Progress(sngNextMark, blnCancel)
            If blnCancel Then
                ' Si el cliente se cansa de esperar
                '   y desea cancelar la tarea,
                '   produce el evento Complete con
                '   True (cancelado).
                RaiseEvent Complete(True)
                Exit Sub
            End If
            sngNextMark = sngNextMark + 0.1!
        End If
    Next
    ' Si se termina con éxito, produce el
    '   evento Complete (no cancelado).
    RaiseEvent Complete(False)
End Sub
