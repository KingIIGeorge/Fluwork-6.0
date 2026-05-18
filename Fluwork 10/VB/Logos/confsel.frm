VERSION 5.00
Begin VB.Form confsel 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2505
   Icon            =   "confsel.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "NO DISPONIBLE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "CONFIRMADO"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "NO CONFIRMADO"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&ACEPTAR"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   2400
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   2400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1080
   End
End
Attribute VB_Name = "confsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Byte
For i = 0 To 2
If Option1(i).Value = True Then
registro.confirmacion = Option1(i).Caption
End If
Next i
Unload Me

End Sub

