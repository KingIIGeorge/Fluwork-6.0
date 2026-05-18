VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CoffeeWatch2"
   ClientHeight    =   3765
   ClientLeft      =   1800
   ClientTop       =   1500
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   4365
   Begin VB.CommandButton cmdMT 
      Caption         =   "Demostración de &múltiples subprocesos"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
   End
   Begin VB.ListBox lstCallBacks 
      Height          =   2235
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdCallBacks 
      Caption         =   "&Comenzar a recibir devolución de llamadas"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdEvents 
      Caption         =   "Comenzar a recibir &eventos CoffeeReady"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox lstEvents 
      Height          =   2235
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Almacenamiento a nivel de módulo para el
'   objeto CoffeeMonitor compartido. La
'   variable se declara WithEvents, para que
'   puedan tratarse los eventos del
'   CoffeeMonitor.
Private WithEvents mwcmnEvents As CoffeeMonitor
Attribute mwcmnEvents.VB_VarHelpID = -1
' Objeto Connector utilizado para obtener una
'   referencia al CoffeeMonitor compartido.
Private mcctEvents As Connector

' Almacenamiento a nivel de módulo para una
'   referencia al objeto CoffeeMonitor2 utilizado
'   en la demostración con método de llamada de retorno.
Private mcm2CallBacks As CoffeeMonitor2
' Almacenamiento para una referencia al objeto de
'   llamada de retorno.
Private mNotifyMe As New NotifyMe

' Botón de comando empleado para activar y
'   desactivar la recepción de eventos.
Private Sub cmdEvents_Click()
    Static blnInUse As Boolean
    
    If blnInUse Then
        ' Estableciendo Nothing en una variable
        '   WithEvents se desconecta al objeto de
        '   sus procedimientos de evento, de forma
        '   que no se reciben más eventos.
        Set mwcmnEvents = Nothing
        '
        ' Fin del objeto Connector.
        Set mcctEvents = Nothing
        cmdEvents.Caption = "Comenzar a recibir eventos CoffeeReady"
    Else
        Set mcctEvents = New Connector
        '
        ' El objeto Connector suministra una referencia
        '   al CoffeeMonitor compartido. Cuando la
        '   la referencia se coloca en la variable
        '   WithEvents, el objeto se conecta a su
        '   procedimiento de evento, de modo que
        '   CoffeeWatch puede empezar a recibir eventos.
        Set mwcmnEvents = mcctEvents.CoffeeMonitor
        cmdEvents.Caption = "PARAR de recibir eventos CoffeeReady"
    End If
    blnInUse = True Xor blnInUse
End Sub

' Botón de comando utilizado para activar y desactivar
'   las llamadas de retorno.
Private Sub cmdCallBacks_Click()
    Static blnInUse As Boolean
    Dim ct2 As New Connector2
    
    If blnInUse Then
        ' Indica a CoffeeMonitor2 que ya no se desean
        '   llamadas de retorno.
        Call mcm2CallBacks.CeaseCallBacks(mNotifyMe)
        '
        ' Libera la copia compartida de CoffeeMonitor2.
        Set mcm2CallBacks = Nothing
        cmdCallBacks.Caption = "Comenzar a recibir devolución de llamadas"
    Else
        ' Obtiene un objeto Connector2, y lo usa para
        '   conseguir una referencia a la copia compartida
        '   de CoffeeMonitor2.
        Set ct2 = New Connector2
        Set mcm2CallBacks = ct2.CoffeeMonitor2
        '
        ' Indica a CoffeeMonitor2 que empiece a hacer
        '   llamadas al objeto NotifyMe (el objeto se
        '   crea aquí implícitamente, ya que la
        '   variable se declara As New).
        Call mcm2CallBacks.TellMeReady(mNotifyMe)
        cmdCallBacks.Caption = "PARAR de recibir devolución de llamadas"
    End If
    blnInUse = True Xor blnInUse
End Sub

Private Sub cmdMT_Click()
    ' Al pasar a la demostración de múltiples
    '   threads, deben desactivarse las llamadas
    '   de retorno y los eventos. Pregunta al
    '   usuario si está de acuerdo. Si la respuesta
    '   es vbNo, continúa y comienza la demostración
    '   de todos modos.
    If (Not mwcmnEvents Is Nothing) Or _
            (Not mcm2CallBacks Is Nothing) Then
        Select Case MsgBox("Es necesario desactivar la notificaciones de eventos y devoluciones de llamadas antes de iniciar la demostración de subprocesos con uso intensivo del procesador. ¿Desea desactivarlas ahora?", _
                vbYesNoCancel, "Comenzar la demostración de subprocesos múltiples")
            Case vbYes
                ' Esto no descarga el formulario,
                '   sólo detiene las demostraciones.
                Call Form_Unload(False)
            Case vbCancel
                Exit Sub
        End Select
    End If
    cmdMT.Enabled = False
    frmThread.Show vbModeless
End Sub

' Cierra las demostraciones en ejecución.
'
Private Sub Form_Unload(Cancel As Integer)
    If Not mwcmnEvents Is Nothing Then
        Call cmdEvents_Click
    End If
    If Not mcm2CallBacks Is Nothing Then
        Call cmdCallBacks_Click
    End If
End Sub

' Cuando el objeto CoffeeMonitor envía un
'   evento CoffeeReady, lo agrega al cuadro
'   de lista. Si hay más de diez elementos,
'   elimina el más antiguo.
'
Private Sub mwcmnEvents_CoffeeReady()
    With lstEvents
        .AddItem Format$(Now, "ddd hh:mm:ss"), 0
        If .ListCount > 10 Then .RemoveItem 10
    End With
End Sub

