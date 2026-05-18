VERSION 5.00
Begin VB.Form frmCYOCC 
   Caption         =   "Crear sus propias clases de colección"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBricks 
      Caption         =   "Casa de ladrillos"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdSticks 
      Caption         =   "Casa de madera"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdStraw 
      Caption         =   "Casa de paja"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmCYOCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmCYOCC muestra sus formularios subsidiarios utilizando
'   las variables globales ocultas que Visual
'   Basic crea para cada clase formulario
'   (en este caso frmStraw, frmSticks y frmBricks).
'   La variable global oculta
'   se describe en "Ciclo de vida
'   de los formularios de Visual Basic" dentro de los Libros en pantalla.
'
' Los formularios se muestran como no modales, y
'   frmCYOCC se especifica así mismo (Me) como
'   propietario de cada uno de ellos.  Como resultado, los
'   formularios aparecen sobre
'   frmCYOCC.

' Casa de paja -- el objeto Collection público.
Private Sub cmdStraw_Click()
    frmStraw.Show vbModeless, Me
End Sub

' Casa de madera -- el objeto Collection privado.
Private Sub cmdSticks_Click()
    frmSticks.Show vbModeless, Me
End Sub

' Casa de ladrillos -- la clase Collection.
Private Sub cmdBricks_Click()
    frmBricks.Show vbModeless, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Si alguno de los formularios no modales
    '   siguen cargados, los cierra.
    If Not frmBricks Is Nothing Then
        Unload frmBricks
        Set frmBricks = Nothing
    End If
    If Not frmSticks Is Nothing Then
        Unload frmBricks
        Set frmBricks = Nothing
    End If
    If Not frmStraw Is Nothing Then
        Unload frmBricks
        Set frmBricks = Nothing
    End If
    
    ' Establece la variable global oculta de frmCYOCC
    '   a Nothing, liberando sus recursos.
    Set frmCYOCC = Nothing
End Sub
