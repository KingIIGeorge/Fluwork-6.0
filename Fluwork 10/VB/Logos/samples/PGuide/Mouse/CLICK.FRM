VERSION 5.00
Begin VB.Form frmClickLine 
   Caption         =   "Clic en una línea"
   ClientHeight    =   4620
   ClientLeft      =   450
   ClientTop       =   2085
   ClientWidth     =   4995
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4620
   ScaleWidth      =   4995
End
Attribute VB_Name = "frmClickLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Line -(X, Y)
End Sub

