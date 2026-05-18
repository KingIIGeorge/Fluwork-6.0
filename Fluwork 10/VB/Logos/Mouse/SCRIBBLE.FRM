VERSION 5.00
Begin VB.Form frmScribble 
   Caption         =   "Garabatos"
   ClientHeight    =   4590
   ClientLeft      =   1125
   ClientTop       =   2115
   ClientWidth     =   4995
   ClipControls    =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4590
   ScaleWidth      =   4995
End
Attribute VB_Name = "frmScribble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim DrawNow As Integer

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawNow = -1
    CurrentX = X
    CurrentY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawNow Then
        Line -(X, Y)
        Circle (X, Y), 50
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawNow = 0
End Sub

