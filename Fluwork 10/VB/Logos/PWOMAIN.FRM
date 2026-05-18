VERSION 5.00
Begin VB.Form frmProgWObMain 
   Caption         =   "Programación con objetos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImplements 
      Caption         =   "Polimorfismo y la palabra clave &Implements"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdEvents 
      Caption         =   "Inicar y controlar &eventos"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton cmdFriends 
      Caption         =   "Pasar UDT entre objetos con funciones &friend"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdCYOCC 
      Caption         =   "&Crear sus propias clases de colección"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "frmProgWObMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmProgWObMain muestra otros formularios que utilizan
'   las variables globales ocultas que Visual
'   Basic crea para cada clase formulario
'   (en este caso frmFriends, frmCYOCC, ______, y _____,).
'  La variable global oculta
'   está descrita en "Ciclo de vida de
'   los formularios de Visual Basic" dentro de los Libros en pantalla.
'
' Los formularios se muestran en forma no modal, y
'   frmProgWObMain se especifica a sí mismo (Me)
'   como propietario de los formularios.  Como resultado,
'   los formularios siempre aparecen
'   sobre frmProgWObMain.

Private Sub cmdEvents_Click()
    frmEvents.Show vbModeless, Me
End Sub

' Miembros Friend que pasan UDTs
Private Sub cmdFriends_Click()
    frmFriends.Show vbModeless, Me
End Sub

' Crear sus propias clases colección
Private Sub cmdCYOCC_Click()
    frmCYOCC.Show vbModeless, Me
End Sub

Private Sub cmdImplements_Click()
    frmImplements.Show vbModeless, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmFriends Is Nothing Then
        Unload frmFriends
        Set frmFriends = Nothing
    End If
    If frmCYOCC Is Nothing Then
        Unload frmCYOCC
        Set frmCYOCC = Nothing
    End If
    If frmEvents Is Nothing Then
        Unload frmEvents
        Set frmEvents = Nothing
    End If
    If frmImplements Is Nothing Then
        Unload frmImplements
        Set frmImplements = Nothing
    End If
End Sub
