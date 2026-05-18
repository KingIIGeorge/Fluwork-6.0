VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoffeeMonitor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' > Encontrará una descripción general de esta aplicación de ejemplo
'   buscando Coffee en la Ayuda en pantalla.
' > AboutCof.Txt, en la carpeta Documentos relacionados de
'   CoffWat2.vbp, también contiene información sobre este ejemplo.

' Clase CoffeeMonitor2
' --------------------
' Al igual que el objeto CoffeeMonitor, CoffeeMonitor2
'   monitoriza una interfaz serie imaginaria con una
'   cafetera de alta tecnología, utilizando un cronómetro
'   para determinar la frecuencia de comprobación del
'   estado del café.
'
' En lugar de provocar un evento cuando el café está listo,
'   el objeto CoffeeMonitor2 invoca un método con llamada de
'   retorno que debe implementar una de las clases del cliente.
'   Este método se declara en la clase ICallBack.
'
' (Como la cafetera de alta tecnología no se ha inventado aún,
'   esta aplicación se limita a invocar el método cada
'   diez segundos).
'
' IMPORTANTE: Para simplificar un ejemplo bastante complejo, se
'   ha incluido la clase ICallBack en este proyecto.
'   Esto NO funcionará en los sistemas reales, que normalmente
'   pasan por varias versiones. Si se realiza una segunda versión
'   de Coffee2, se incrementará el número de versión de la ID
'   de interfaz, y los clientes anteriores no funcionarán con la
'   nueva versión de Coffee2. Las interfaces estándar, como
'   ICallBack deben crearse por sí mismas en pequeñas DLL que
'   puedan referenciar tanto el cliente como el componente.
'   Una vez que hay aplicaciones terminadas que utilizan una
'   interfaz, ésta nunca debe cambiarse. Si desea más información,
'   busque "polimorfismo" en los Libros en pantalla.
'
' Observe que el valor de la propiedad Instancing de la clase
'   CoffeeMonitor2 es PublicNotCreatable. Esto significa que los
'   clientes no pueden crear un CoffeeMonitor2, sino sólo obtener
'   una referencia al CoffeeMonitor2 compartido creando un objeto
'   Connector2 y accediendo a su propiedad CoffeeMonitor2.
'
' Al igual que CoffeeMonitor, la clase CoffeeMonitor2 corrige
'   el error descrito en el tema "Uso de la clase CoffeeMonitor
'   compartida", en "Creación de un componente EXE ActiveX" de los
'   Libros en pantalla, y según el cual en ocasiones era posible
'   crear varios objetos CoffeeMonitor.

' ===================================================================
'  ¡ADVERTENCIA! Los cronómetros sólo de código son inherentemente
'       peligrosos en el entorno de desarrollo de Visual Basic, ya
'       que el sistema retorna ciegamente al código hasta que se
'       desactive el cronómetro con una llamada de API. Es más
'       seguro utilizar controles Timer durante la mayoría del
'       proceso de desarrollo, y pasar a los cronómetros con llamada
'       de retorno muy al final.
' ===================================================================

Const ICN_ARRAYINCREMENT = 10

' maicnClients almacena referencias a todos los clientes que han
' ------------   solicitado retornos. (Observe que el proceso es
'   distinto del uso de eventos en CoffeeMonitor; un evento puede
'   recibirlo cualquier número de clientes, mientras que las
'   llamadas de retorno deben hacerse una a una). Se utiliza una
'   matriz, en lugar de una colección, debido a que los objetos
'   Collection almacenan los objetos con tipo Variant, lo que
'   implica su vinculación tardía.
Private maicnClients() As ICoffeeNotify
Private mlngMaxClients As Long

' mwXTimer contiene una referencia a un cronómetro de sólo código
' -------   que indica a CoffeeMonitor2 cuándo comprobar la cafetera.
'   Como la variable se declara WithEvents, el objeto CoffeeMonitor2
'   recibe los eventos Tick del objeto XTimer. (vea Sub mwXTimer_Tick,
'   más abajo. El código del objeto XTimer se encuentra en XTimers.vbp.
Private WithEvents mwXTimer As XTimer
Attribute mwXTimer.VB_VarHelpID = -1

Private Sub Class_Initialize()
    ' Asigna espacio en la matriz de objetos cliente.
    mlngMaxClients = ICN_ARRAYINCREMENT
    ReDim maicnClients(1 To mlngMaxClients)
    '
    ' Crea el objeto XTimer. Cuando se hace esta asignación, Visual
    '   Basic conecta el  evento Tick del XTimer al procedimiento de
    '   evento mwXTimer_Tick (ver más abajo).
    Set mwXTimer = New XTimer
    '
    ' El cronómetro se establece de modo que se dispare cada diez
    '   segundos (10.000 milisegundos).
    mwXTimer.Interval = 10000
    mwXTimer.Enabled = True
End Sub

Private Sub Class_Terminate()
    Dim intCt As Integer
    
    ' Es importante desactivar el XTimer antes de liberarlo.
    '   Como se describe en XTimers.vbp, abandonar un XTimer
    '   en ejecución supone en esencia perder un cronómetro
    '   del sistema hasta que se cierra finalmente XTimers.DLL.
    mwXTimer.Enabled = False
    Set mwXTimer = Nothing
    '
    ' Libera todos los clientes con llamada de retorno que queden,
    '   en caso de que ellos hayan liberado CoffeeMonitor2 sin
    '   solicitar primero que terminen las notificaciones.
    For intCt = 1 To mlngMaxClients
        Set maicnClients(intCt) = Nothing
    Next
    '
    Debug.Print "Devolución de llamadas de CoffeMonitor2 terminadas a las " & Now
End Sub

' TellMeReady es llamado por los clientes que desean recibir una llamada de
' -----------   retorno cuando el café esté listo. El cliente debe
'   implementar la interfaz ICoffeeNotify, definida en la clase
'   ICoffeeNotify.
'
Public Sub TellMeReady(ByVal icn As ICoffeeNotify)
    Dim lngCt As Long
    
    ' Busca un lugar en la matriz de interfaces.
    For lngCt = 1 To mlngMaxClients
        If maicnClients(lngCt) Is Nothing Then Exit For
    Next
    '
    ' Si no lo encuentra, aumenta la matriz.
    If lngCt > mlngMaxClients Then
        mlngMaxClients = mlngMaxClients + ICN_ARRAYINCREMENT
        ReDim Preserve maicnClients(1 To mlngMaxClients)
    End If
    '
    Set maicnClients(lngCt) = icn
    '
    ' Da al objeto el índice de su entrada, como clave para
    '   una búsqueda rápida cuando se solicite la desconexión.
    icn.NotifyID = lngCt
End Sub

' CeaseCallBacks retira al cliente de la lista de objetos que
' --------------   reciben notificaciones de llamada de retorno,
'   usando la clave asignada al objeto cuando éste solicitó las
'   notificaciones.
'
Public Sub CeaseCallBacks(ByVal icn As ICoffeeNotify)
    Set maicnClients(icn.NotifyID) = Nothing
End Sub

' mwXTimer_Tick es el procedimiento de evento que usa CoffeeMonitor2
' -------------   para recibir los eventos Tick del objeto XTimer.
'   El nombre de un procedimiento de evento asociado a una variable
'   WithEvents siempre tiene como prefijo el nombre de la variable.
'
Private Sub mwXTimer_Tick()
    Dim lngCt As Long
    
    ' (Se omite el código para comprobar el puerto serie).
    '
    On Error Resume Next
    '
    ' El método con llamada de retorno debe invocarse para
    '   cada objeto que haya solicitado una notificación.
    For lngCt = 1 To mlngMaxClients
        If Not maicnClients(lngCt) Is Nothing Then
            maicnClients(lngCt).CoffeeReady
            If Err.Number <> 0 Then
                ' El error &H80010005 se ignora, ya que puede
                '   deberse a que el objeto cliente no responda
                '   temporalmente.
                If Err.Number <> &H80010005 Then
                    ' Si una aplicación cliente se ha cerrado sin
                    '   terminar las notificaciones, la retira de
                    '   la lista.
                    Set maicnClients(lngCt) = Nothing
                End If
                '
                ' Cuando se usaOn Error Resume Next,
                '   es necesario borrar el número de error
                '   después de cada error.
                Err.Number = 0
            End If
        End If
    Next
End Sub

