VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de FluWork "
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Editado por Jorgito 18-5-26"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      Caption         =   "Fichas Indexadas."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   480
      X2              =   5040
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Copyright(C) 1992-2000 DarkSoft Development "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   " Windows Compatible "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Fluwork 10 D3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "about.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   6375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim canrdefichas As Long
cantdefichas = 0
cantdefichas = getlastfichanumber - BASE
Label8.Caption = cantdefichas

End Sub

