VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Widget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' >> Se ve mejor en la vista de Módulo completo. <<
'
' Almacenamiento para el número de ID de depuración.
Private mlngDebugID As Long
Implements IDebug

' Al definir los números de error en una Enum pública
'   se les hace visibles en todo
'   el proyecto.
Public Enum WidgetErrors
    wdgERRTaskCanceled = 1059
End Enum

' El evento PercentDone se desencadena periódicamente
'   durante LongTask, para notificar al llamador
'   sobre su progreso. Los argumentos del evento son
'   el porcentaje de ejecución y un argumento Cancel
'   por referencia que el llamador puede establecer a
'   True para cancelar LongTask.
Event PercentDone(ByVal Percent As Double, _
    Cancel As Boolean)
    
' El método LongTask simula una tarea de larga
'   duración que produce
'   el evento PercentDone y permite que el llamador
'   cancele la operación.
'
' El primer argumento indica a LongTask cuánto
'   quiere que dure la tarea simulada.
'   El segundo argumento es el mínimo
'   intervalo entre los eventos producidos
'   para notificar acerca de su progreso.
'
' El uso de un intervalo de tiempo para determinar
'   cuando producir el evento ofrece
'   resultados más consistentes en diferentes
'   equipos. Puede ver otra
'   solución en LongTask2.
'
Public Sub LongTask(ByVal Duration As Double, _
        ByVal MinimumInterval As Double)
    Dim dblThreshold As Double
    Dim dblStart As Double
    Dim blnCancel As Boolean
    
    dblStart = Timer
    dblThreshold = MinimumInterval
    
    Do While Timer < (dblStart + Duration)
        ' En una aplicación real, aquí se haría
        '   una unidad de trabajo. El trabajo se
        '   dividiría de forma que las unidades
        '   no fueran demasiado grandes
        '   (mucho tiempo entre notificaciones)
        '   ni demasiado pequeñas (cuantas más
        '   notificaciones, menos eficiente será la
        '   tarea LongTask).
        
        ' Después de cada unidad de trabajo, verifica si
        '   es tiempo de notificar al
        '   llamador acerca del progreso de LongTask.
        If Timer > (dblStart + dblThreshold) Then
            ' Desencadena el evento; ¡la ejecución de
            '   LongTask no continuará
            '   hasta que termine el
            '   procedimiento de evento del llamador!
            RaiseEvent PercentDone( _
                dblThreshold / Duration, _
                blnCancel)
            '
            ' Comprueba si el llamador
            '   quiere cancelar LongTask.
            If blnCancel Then
                Err.Raise vbObjectError + wdgERRTaskCanceled, , _
                    "Tarea cancelada"
' NOTA: Si su programa se detiene aquí, haga clic con el botón
'   secundario para presentar el menú de contexto de la ventana de código.
'   Haga clic en Alternar, después haga clic en Interrupción en errores no
'   controlables. Presione F5 para seguir ejecutando el
'   programa. (Puede que tenga que presionar Alt+Tab para
'   pasar del formulario Events), aquí se explica
'   porqué se cambia la configuración:
' La configuración predeterminada, Parada en módulo de clase,
'   es útil si tiene un error en un
'   método de una clase, puesto que permite que
'   Visual Basic se detenga DENTRO del módulo de clase,
'   en el punto del error. Si su clase
'   produce errores de forma rutinaria, como aquí, esto no es
'   tan conveniente.
' Puede establecer que la opción predeterminada sea Interrupción en errores
'   no controlados, en la ficha General del cuadro de diálogo Opciones,
'   accesible desde el menú Herramientas.
'   Si lo hace, recuerde que cuando se detiene
'   en una llamada a un método y quiere ir al
'   punto del error, puede utilizar el menú de contexto
'   de la venta de código para cambiar a
'   Parada en módulo de clase.
' Observe que también puede utilizar Alt+F5 para saltar
'   un error cuando se encuentra en el modo Parada en
'   módulo de clase (o Alt+F8 para seguir). Si estas teclas
'   le llevan a la misma línea de código,
'   es que no hay un controlador de errores disponible.
' Para obtener más información, vea "Depuración de los módulos de clase"
'   en los Libros en pantalla.
' [End Digression]
            End If
            '
            ' Establece el umbral de la siguiente
            '   notificación.
            dblThreshold = dblThreshold + MinimumInterval
       End If
    Loop
End Sub
 
' LongTask2 también simula una tarea de larga
'   duración que produce el
'   evento PercentDone, y permite que el llamador
'   cancele la operación.
'
' La tarea simulada consiste en cálculos
'   de coma flotante repetitivos. El primer
'   argumento indica a LongTask2 cuántas
'   iteraciones quiere que tenga la tarea.
'   El segundo argumento indica la cantidad
'   de porcentaje de ejecución que produce el
'   evento de notificación. Observe que
'   este método produce notificaciones
'   a intervalos de tiempo diferentes
'   una variación que puede ser debida a
'   las diferencias en el rendimiento de la máquina.
'
' En contraste, LongTask utiliza un
'   intervalo de tiempo para determinar la frecuencia
'   con la que producir el evento; esto produce
'   resultados más consistentes en equipos
'   diferentes.
'
Public Sub LongTask2(ByVal Iterations As Long, _
        ByVal PercentChange As Byte)
    Dim lngThreshold As Long
    Dim dblIterationsPerEvent As Double
    Dim lngCt As Long
    Dim dblDummy As Double
    Dim blnCancel As Boolean
    
    dblIterationsPerEvent = Iterations _
        * (CDbl(PercentChange) / 100)
    lngThreshold = dblIterationsPerEvent
    
    For lngCt = 1 To Iterations
        ' En una aplicación real, aquí se haría
        '   una unidad de trabajo. El trabajo se
        '   dividiría de forma que las unidades
        '   no fueran demasiado grandes
        '   (mucho tiempo entre notificaciones)
        '   ni demasiado pequeñas (cuantas más
        '   notificaciones, menos eficiente será la
        '   tarea LongTask).
       dblDummy = 3.14159 * 2640 * 2640
        
        ' Después de cada unidad de trabajo, verifica si
        '   es tiempo de notificar al
        '   llamador acerca del progreso de LongTask.
        If lngCt > lngThreshold Then
            ' Desencadena el evento; ¡la ejecución de
            '   LongTask no continuará
            '   hasta que termine el
            '   procedimiento de evento del llamador!
            RaiseEvent PercentDone( _
               lngCt * 100 / Iterations, _
                blnCancel)
            '
            ' Comprueba si el llamador
            '   quiere cancelar LongTask.
            If blnCancel Then
                Err.Raise vbObjectError + wdgERRTaskCanceled, , _
                    "Tarea cancelada"
            End If
            '
            ' Establece el umbral de la siguiente
            '   notificación.
            lngThreshold = lngThreshold + dblIterationsPerEvent
        End If
    Next
End Sub

Private Sub Class_Initialize()
    mlngDebugID = DebugInit(Me)
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


