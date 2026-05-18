VERSION 5.00
Begin VB.Form frmReceiver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje recibido por el controlador de eventos de difusión múltiple"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkReceive 
      Caption         =   "&Recibir mensajes"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.CheckBox chkGarble 
      Caption         =   "&Ignorar mensajes"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblMessage 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' La variable Source se declara utilizando la
'   palabra clave WithEvents, de forma que
'   los eventos del objeto frmEvents
'   asignados a ella puedan ser gestionados (vea el
'   procedimiento de evento Source_Broadcast).  Cuando se
'   asigna la referencia, Visual Basic
'   conecta el objeto con el procedimiento
'   de evento.
Private WithEvents Source As frmEvents
Attribute Source.VB_VarHelpID = -1

Private Sub chkReceive_Click()
    If chkReceive = vbChecked Then
        Set Source = frmEvents
    Else
        ' Al establecer la variable WithEvents
        '   a Nothing se desconecta el
        '   procedimiento de evento del
        '   objeto frmEvents.
        Set Source = Nothing
    End If
End Sub

Private Sub Form_Load()
    chkReceive = vbChecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Source = Nothing
End Sub

' Procedimiento de evento del evento Broadcast.
'   Puesto que el argumento Message es ByRef,
'   puede ser alterado y dichos cambios
'   serán visibles por los siguientes
'   receptores del evento.
Private Sub Source_Broadcast(Message As String)
    Dim intCt As Integer
    lblMessage = Message
    '
    ' Si Ignorar mensajes está activado, distorsiona el
    '   mensaje.
    If chkGarble = vbChecked Then
        ' Utiliza un paso aleatorio que toca cada
        '   carácter entre el tercero (mínimo) y
        '   el noveno (máximo).
        For intCt = 1 To Len(Message) Step (Int(7 * Rnd) + 3)
            ' En los caracteres que van a ser distorsionados,
            '   se cambia su valor ASCII en una cantidad aleatoria
            '   entre -3 y +3.
            Mid$(Message, intCt, 1) = Chr$(Asc(Mid$(Message, intCt, 1)) + Int(7 * Rnd) - 3)
        Next
    End If
End Sub
