VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoffeeTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CoffeeTracker es un pequeño objeto privado
' =============     para hacer un seguimiento
'   de una tarea que ejecute un objeto Coffee.
'   La razón de utilizar un objeto es que las
'   variables WithEvents no pueden estar
'   contenidas en una matriz. Para tener un
'   número variable de ellas, es necesario
'   tener una clase de objetos que las contengan.
'
' Establezca ThreadID y Size antes de empezar
'   la tarea larga; y la propiedad Coffee
'   INMEDIATAMENTE antes de llamar a StartLongTask.
'   El procedimiento NewTracker asigna el ID en
'   frmThread; es el índice del objeto en el
'   objeto Collection CoffeeTrackers.

Public ThreadID As Long
Public Size As Long
Public ID As String

' Almacenamiento para el objeto Coffee observado.
Private WithEvents mwCoffee As Coffee
Attribute mwCoffee.VB_VarHelpID = -1
'
' Hora de inicio (del API timeGetTimer).
Private mlngStart As Long

Public Property Get Coffee() As Coffee
    Set Coffee = mwCoffee
End Property
Public Property Set Coffee(ByVal NewValue As Coffee)
    ' Guarda la hora de inicio.
    mlngStart = timeGetTime
    Set mwCoffee = NewValue
End Property

' El objeto Coffee produce un evento Complete
'   cuando la tarea observada se completa.
'   CoffeeTracker coloca información sobre
'   la tarea (ID de subroceso, tamaño y segundos
'   por iteración) en un cuadro de lista de
'   frmThread.
'
Private Sub mwCoffee_Complete(ByVal Canceled As Boolean)
    Dim lngEnd As Long
    Dim dblElapsed As Double
    
    lngEnd = timeGetTime
    '
    ' Libera el objeto Coffee.
    Set mwCoffee = Nothing
    '
    ' Agrega una línea de informe al cuadro de lista.
    If Canceled Then
        frmThread.lstResults.AddItem ThreadID _
            & " (" & Size & ") cancelado", 0
    Else
        frmThread.lstResults.AddItem ThreadID _
            & " (" & Size & ") " _
            & (CDbl(lngEnd) - mlngStart) / Size / 1000# _
            & " seg/iteración", 0
    End If
    '
    ' CoffeeTracker elimina su referencia
    '   de la colección, quedando él mismo
    '   sin referencias, por lo que puede
    '   terminarse.
    frmThread.CoffeeTrackers.Remove ID
End Sub

' En el caso de tareas prolongadas, los eventos
'   Coffee producen un evento Progress para cada
'   10% de la tarea completa. CoffeeTracker agrega
'   una entrada al cuadro de lista de frmThread.
'
Private Sub mwCoffee_Progress(ByVal PercentDone As Single, Cancel As Boolean)
    frmThread.lstResults.AddItem ThreadID _
        & " (" & Size & ") " _
        & Format$(PercentDone * 100, "#0.0") & "%", 0
    '
    ' Cada CoffeeTracker que observa el
    '   indicador global Cancel desactiva
    '   la tarea prolongada que ha estado
    '   observando. El objeto Coffee producirá
    '   entonces el evento Complete.
    '   (ver más arriba).
    If frmThread.CancelAll Then Cancel = True
End Sub

