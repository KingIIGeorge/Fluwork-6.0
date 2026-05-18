VERSION 5.00
Begin VB.Form frmEvents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iniciar y controlar eventos"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddForm 
      Caption         =   "Agregar &receptor"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   4455
   End
   Begin VB.CommandButton cmdPercentDone 
      Caption         =   "&Iniciar una tarea larga que usa un evento para informar de su ejecución"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Mensaje después de que todos los recipientes lo han controlado:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2595
   End
   Begin VB.Label lblPercentDone 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lblEcho 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   $"PWOEvent.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================
'     Declaraciones para la demostración de Broadcast
'
' El evento Broadcast tiene un argumento,
'   el mensaje que se va a enviar.  El argumento
'   es ByRef, de forma que los receptores pueden modificarlo.
Event Broadcast(Message As String)

' Colección de receptores.
Private mcolReceivers As New Collection

' =============================================================
'       Declaraciones y código para la demostración de
'           PercentDone
'
Private WithEvents mWidget As Widget
Attribute mWidget.VB_VarHelpID = -1
Private mblnCancel As Boolean

Private Sub mWidget_PercentDone(ByVal Percent As Double, Cancel As Boolean)
    lblPercentDone.Caption = CInt(100 * Percent) & " % completado"
    DoEvents
    If mblnCancel Then Cancel = True
End Sub

Private Sub cmdPercentDone_Click()
    Static blnProcessing As Boolean
    If blnProcessing Then
        mblnCancel = True
    Else
        blnProcessing = True
        cmdPercentDone.Caption = "&Cancelar tarea"
        mblnCancel = False
        lblPercentDone.Caption = "0 % completado"
        lblPercentDone.Refresh
        
        ' Crea un objeto Widget e inicia la
        '   tarea de larga duración.
        Set mWidget = New Widget
        On Error Resume Next
        Call mWidget.LongTask(14.4, 0.9)
        '
        ' Mira si la llamada ha terminado porque ha
        '   sido cancelada (no se puede leer
        '   mblnCancel para saberlo, puesto que
        '   podría haber sido establecida al volver de
        '   LongTask).
        If Err.Number = 0 Then
            lblPercentDone.Caption = "Tarea terminada"
        ElseIf Err.Number = vbObjectError + wdgERRTaskCanceled Then
            lblPercentDone.Caption = "Tarea cancelada"
        Else
            ' (Se omite el control del resto de los errores).
            lblPercentDone.Caption = "Sucedió algo malo"
        End If
        Set mWidget = Nothing
        cmdPercentDone.Caption = "&Iniciar una tarea larga que usa un evento para informar de su ejecución"
        blnProcessing = False
    End If
End Sub

' =============================================================
'        Código de la demostración de Broadcast
'
Private Sub cmdAddForm_Click()
    Dim frm As New frmReceiver
    ' Seguimiento de los receptores.
    mcolReceivers.Add frm
    frm.Show vbModeless, Me
    Me.SetFocus
    txtMessage.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As frmReceiver
    On Error Resume Next
    Do While mcolReceivers.Count > 0
        Unload mcolReceivers(1)
        mcolReceivers.Remove 1
    Loop
End Sub

Private Sub txtMessage_Change()
    Dim strMessage As String
    strMessage = txtMessage.Text
    '
    ' Desencadena el evento Broadcast.  Observe que
    '   no hay forma de saber si hay
    '   receptores que controlen el
    '   evento.
    RaiseEvent Broadcast(strMessage)
    '
    ' Presenta el mensaje después de que todos
    '   los receptores (si hay alguno) lo hayan controlado.
    '   Observe que no hay manera de saber qué
    '   receptor alteró el mensaje,
    '   o qué valores internos
    '   el mensaje ha podido tener.
    lblEcho = strMessage
End Sub
