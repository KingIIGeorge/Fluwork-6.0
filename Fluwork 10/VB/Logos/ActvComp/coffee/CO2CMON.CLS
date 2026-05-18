VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoffeeMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' > Encontrará una descripción general de esta aplicación de
'   ejemplo buscando Coffee en la Ayuda en pantalla.
' > AboutCof.Txt, en la carpeta Documentos relacionados de
'   CoffWat2.vbp, también contiene información sobre el ejemplo.

' La clase CoffeeMonitor monitoriza una interfaz serie imaginaria
' ----------------------   con una cafetera de alta tecnología,
'   empleando un cronómetro con llamada de retorno (objeto XTimer)
'   para determinar la frecuencia de comprobación del estado del
'   café. Cuando el café está listo, el objeto CoffeeMonitor
'   produce un evento para notificarlo a los clientes. (Como la
'   cafetera de alta tecnología aún no se ha inventado, esta
'   aplicación de ejemplo se limita a producir el evento CoffeeReady
'   cada diez segundos).
'
' (Vea también la clase CoffeeMonitor2, que ilustra la notificación
'   con un método de llamada de retorno, en lugar de con un evento).
'
' Observe que la propiedad Instancing de la clase CoffeeMonitor tiene
'   el valor PublicNotCreatable. Esto significa que los clientes no
'   pueden crear CoffeeMonitors; sino sólo obtener una referencia al
'   CoffeeMonitor compartido creando un objeto Connector y accediendo
'   a su propiedad CoffeeMonitor.
'
' Con un cronómetro de código únicamente, esta versión de CoffeeMonitor
'   soluciona el error descrito en el tema "Uso de la clase CoffeeMonitor
'   compartida" en "Creación de un componente EXE ActiveX", de los manuales
'   en pantalla, donde en ocasiones era posible crear varios objetos
'   CoffeeMonitor.
'
' La causa de tal error era el código de 'cuenta de uso' que empleaba
'   el objeto Connector para liberar la referencia global al
'   CoffeeMonitor compartido. La cuenta de uso era necesaria debido a
'   que en los procedimientos paso a paso de los Libros en pantalla
'   se utilizaba un formulario oculto para contener un control Timer, y
'   los formularios ocultos impiden que se descarguen los componentes
'   (tal y como se describe en "Iniciar y terminar un componente", en
'   los Libros en pantalla). Esto impedía liberar memoria global,
'   lo que a su vez hacía imposible establecer Nothing en
'   gCoffeeMonitor.
'
' A diferencia del formulario oculto con su control Timer, el objeto
'   XTimer no impide que el componente Coffee2 se descargue cuando el
'   último cliente libere su última referencia. (Observe, sin embargo,
'   que la variable global gCoffeeMonitor mantendrá vivo el objeto
'   CoffeeMonitor hasta que todos los objetos que proporciona Coffee2
'   sean liberados por sus clientes respectivos).

' =================================================================
'  ¡ADVERTENCIA! Los cronómetros sólo de código son inherentemente
'       peligrosos en el entorno de desarrollo de Visual Basic, ya
'       que el sistema retornará ciegamente al código hasta que se
'       desactive el cronómetro con una llamada del API. Es más
'       seguro utilizar controles Timer durante la mayoría del
'       proceso de desarrollo, y pasar a los cronómetros de
'       llamada de retorno muy al final.
' =================================================================


' mwXTimer contiene una referencia a un cronómetro sólo de código que
' -------   indica a CoffeeMonitor cuándo debe comprobar la cafetera.
'   Como la variable se declara con WithEvents, el objeto
'   CoffeeMonitor recibe los eventos Tick del objeto XTimer
'   (vea Sub mwXTimer_Tick, más abajo). Encontrará el código del
'   objeto XTimer en XTimers.vbp.
Private WithEvents mwXTimer As XTimer
Attribute mwXTimer.VB_VarHelpID = -1

' CoffeeReady es el evento que produce CoffeeMonitor para notificar a
' -----------   sus clientes que el café está listo. La palabra clave
'   Public se omite de la declaración del evento debido a que los
'   eventos son siempre públicos.
Event CoffeeReady()
    
Private Sub Class_Initialize()
    ' Lo primero que hace un objeto CoffeeMonitor es crear
    '   el objeto XTimer. Una vez hecha esta asignación,
    '   Visual Basic conecta el evento Tick de XTimer al
    '   procedimiento de evento mwXTimer_Tick (vea más abajo).
    '
    Set mwXTimer = New XTimer
    '
    ' El cronómetro cuenta cada diez segundos (10.000
    '   milisegundos).
    mwXTimer.Interval = 10000
    mwXTimer.Enabled = True
End Sub

Private Sub Class_Terminate()
    ' Es importante desactivar el XTimer antes de
    '   liberarlo. Como se describe en XTimers.vbp, si se
    '   abandona un XTimer en ejecución, en esencia se
    '   pierde un cronómetro del sistema hasta que se
    '   cierre finalmente XTimers.DLL.
    mwXTimer.Enabled = False
    Set mwXTimer = Nothing
    '
    Debug.Print "Eventos de CoffeeMonitor terminados a las " & Now
End Sub

' mwXTimer_Tick es el procedimiento de evento que usa
' -------------   CoffeeMonitor para recibir los eventos
'   Tick del objeto XTimer. El nombre de un procedimiento
'   de evento asociado a una variable WithEvents tiene
'   siempre el nombre de ésta como prefijo.
'
Private Sub mwXTimer_Tick()
    ' (Se omite el código para comprobar el puerto serie).
    '
    ' Notifica al cliente.
    RaiseEvent CoffeeReady
End Sub
