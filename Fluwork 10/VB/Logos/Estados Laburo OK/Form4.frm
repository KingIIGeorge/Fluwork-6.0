VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libreta de Direcciones"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form4.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line4 
      X1              =   5040
      X2              =   5040
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   5040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Archivo creado en:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Finalizo la exportacion de la libreta de direcciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = nombrecompleto
End Sub

