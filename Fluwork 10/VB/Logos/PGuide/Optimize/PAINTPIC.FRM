VERSION 5.00
Begin VB.Form frmPaintPic 
   Caption         =   "Demostración de dibujo con Picture"
   ClientHeight    =   1845
   ClientLeft      =   1710
   ClientTop       =   2355
   ClientWidth     =   4620
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
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1845
   ScaleWidth      =   4620
   Tag             =   "Apparent"
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2580
      Top             =   1305
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Comenzar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3180
      TabIndex        =   1
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1740
      Top             =   1305
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   5
      Left            =   615
      Picture         =   "PaintPic.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   4
      Left            =   555
      Picture         =   "PaintPic.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   405
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   3
      Left            =   480
      Picture         =   "PaintPic.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   330
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   2
      Left            =   405
      Picture         =   "PaintPic.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   1
      Left            =   315
      Picture         =   "PaintPic.frx":1108
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   165
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   240
      Picture         =   "PaintPic.frx":154A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Comenzar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Image imgMoon 
      Height          =   510
      Left            =   3525
      Top             =   300
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asignar imágenes"
      Height          =   195
      Index           =   1
      Left            =   3030
      TabIndex        =   9
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dibujar imágenes"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   1020
      Width           =   1485
   End
End
Attribute VB_Name = "frmPaintPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAction_Click(Index As Integer)
Dim iMaxHeight As Integer, iMaxWidth As Integer
Dim ipicHeight As Integer, iPicWidth As Integer
Dim ipicLeft As Integer, ipicTop As Integer, i As Integer

If cmdAction(Index).Caption = "&Comenzar" Then
  If Index = 0 Then
    Timer1.Interval = 20
    Timer1.Enabled = True
  Else
    Timer2.Interval = 20
    Timer2.Enabled = True
  End If
  cmdAction(Index).Caption = "&Parar"
Else
  If Index = 0 Then
    Timer1.Enabled = False
  Else
    Timer2.Enabled = False
  End If
  cmdAction(Index).Caption = "&Comenzar"
End If

End Sub

Private Sub Form_Load()
  PosForm Me
  Me.Show
End Sub


Private Sub Timer1_Timer()
  Static idx As Integer
  
  frmPaintPic.PaintPicture picMoon(idx), 400, 300
  idx = idx + 1
  If idx > 5 Then idx = 0

End Sub


Private Sub Timer2_Timer()
  Static idx As Integer
  
  imgMoon.Picture = picMoon(idx).Picture
  idx = idx + 1
  If idx > 5 Then idx = 0

End Sub
