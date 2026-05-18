VERSION 5.00
Begin VB.Form trabajos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8055
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "Copperplate Gothic Light"
      Size            =   9.75
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmwdone.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11985
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   6960
      Width           =   9855
   End
   Begin VB.TextBox btl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10560
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
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
      Height          =   375
      Left            =   10320
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   10320
      TabIndex        =   19
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   10320
      TabIndex        =   17
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   10320
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   10320
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   10320
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   10320
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   10320
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   10320
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   10320
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   9
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   6120
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   8
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   5520
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   7
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4920
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4320
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3120
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   9375
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   9375
   End
   Begin VB.TextBox importe 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   10320
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox trabajo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   8640
      TabIndex        =   44
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000008&
      Caption         =   "bytes libres:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6960
      TabIndex        =   43
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000A&
      X1              =   11760
      X2              =   11760
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00E0E0E0&
      X1              =   9960
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00E0E0E0&
      X1              =   9960
      X2              =   11760
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000A&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   9960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   9960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000012&
      Caption         =   "8."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   42
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000012&
      Caption         =   "9."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000012&
      Caption         =   "10."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000012&
      Caption         =   "7."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000012&
      Caption         =   "4."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000012&
      Caption         =   "5."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "6."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "3."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      Caption         =   "2."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000012&
      Caption         =   "1."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000015&
      Caption         =   "importe"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   10440
      TabIndex        =   32
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000008&
      Caption         =   "trabajos realizados"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2520
      TabIndex        =   30
      Top             =   7560
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "modelo:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000004&
      X1              =   120
      X2              =   9960
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      X1              =   120
      X2              =   120
      Y1              =   6960
      Y2              =   7920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   9960
      X2              =   9960
      Y1              =   6960
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      BorderStyle     =   6  'Inside Solid
      X1              =   120
      X2              =   9960
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8280
      TabIndex        =   27
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "telefono:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6960
      TabIndex        =   26
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   7320
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "nombre y apellido:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "numero de ficha:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   7080
      Width           =   2295
   End
End
Attribute VB_Name = "trabajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim BLOQUE As String
Dim i As Byte
BLOQUE = ""
For i = 0 To 8
BLOQUE = BLOQUE & trabajo(i).Text & "|" & importe(i).Text & "|"
Next i
BLOQUE = BLOQUE & trabajo(9).Text & "|" & importe(9).Text
Form1.tsolucion.Text = BLOQUE
Unload Me
End Sub

Private Sub Form_Activate()
Dim BLOQUE As String * 1024
Dim dpos(1 To 19) As Long
Dim i As Long
Dim z As Byte
Dim lastone As Long

BLOQUE = Form1.tsolucion.Text
If tmpficha = 0 Then
frame2.Visible = True
Else
frame2.Visible = False
End If
Label4.Caption = tmpficha
Label6.Caption = registro.fullname
Label8.Caption = registro.telefono
Label3.Caption = registro.modelo
lastone = 1
For z = 1 To 19
For i = lastone To 1024
If Mid$(BLOQUE, i, 1) = "|" Then
lastone = i + 1
dpos(z) = i
Exit For
End If
Next i
Next z

If (dpos(1) = 0) Then
BLOQUE = "|||||||||||||||||||"
For i = 1 To 19
dpos(i) = i
Next i
End If

trabajo(0).Text = Trim(Mid$(BLOQUE, 1, dpos(1) - 1))
importe(0).Text = Trim(Mid$(BLOQUE, dpos(1) + 1, dpos(2) - dpos(1) - 1))
trabajo(1).Text = Trim(Mid$(BLOQUE, dpos(2) + 1, dpos(3) - dpos(2) - 1))
importe(1).Text = Trim(Mid$(BLOQUE, dpos(3) + 1, dpos(4) - dpos(3) - 1))
trabajo(2).Text = Trim(Mid$(BLOQUE, dpos(4) + 1, dpos(5) - dpos(4) - 1))
importe(2).Text = Trim(Mid$(BLOQUE, dpos(5) + 1, dpos(6) - dpos(5) - 1))
trabajo(3).Text = Trim(Mid$(BLOQUE, dpos(6) + 1, dpos(7) - dpos(6) - 1))
importe(3).Text = Trim(Mid$(BLOQUE, dpos(7) + 1, dpos(8) - dpos(7) - 1))
trabajo(4).Text = Trim(Mid$(BLOQUE, dpos(8) + 1, dpos(9) - dpos(8) - 1))
importe(4).Text = Trim(Mid$(BLOQUE, dpos(9) + 1, dpos(10) - dpos(9) - 1))
trabajo(5).Text = Trim(Mid$(BLOQUE, dpos(10) + 1, dpos(11) - dpos(10) - 1))
importe(5).Text = Trim(Mid$(BLOQUE, dpos(11) + 1, dpos(12) - dpos(11) - 1))
trabajo(6).Text = Trim(Mid$(BLOQUE, dpos(12) + 1, dpos(13) - dpos(12) - 1))
importe(6).Text = Trim(Mid$(BLOQUE, dpos(13) + 1, dpos(14) - dpos(13) - 1))
trabajo(7).Text = Trim(Mid$(BLOQUE, dpos(14) + 1, dpos(15) - dpos(14) - 1))
importe(7).Text = Trim(Mid$(BLOQUE, dpos(15) + 1, dpos(16) - dpos(15) - 1))
trabajo(8).Text = Trim(Mid$(BLOQUE, dpos(16) + 1, dpos(17) - dpos(16) - 1))
importe(8).Text = Trim(Mid$(BLOQUE, dpos(17) + 1, dpos(18) - dpos(17) - 1))
trabajo(9).Text = Trim(Mid$(BLOQUE, dpos(18) + 1, dpos(19) - dpos(18) - 1))
importe(9).Text = Trim(Mid$(BLOQUE, dpos(19) + 1))
End Sub

Private Sub Form_Load()
Me.WindowState = 0
Me.Height = 8500
Me.Width = 12075
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Command1.Value = True
End Sub

Private Sub importe_GotFocus(Index As Integer)
Select Case Index
       Case 0
        importe(0).BackColor = QBColor(7)
        trabajo(0).BackColor = QBColor(7)
        Case 1
        importe(1).BackColor = QBColor(7)
        trabajo(1).BackColor = QBColor(7)
        Case 2
        importe(2).BackColor = QBColor(7)
        trabajo(2).BackColor = QBColor(7)
        Case 3
        importe(3).BackColor = QBColor(7)
        trabajo(3).BackColor = QBColor(7)
        Case 4
        importe(4).BackColor = QBColor(7)
        trabajo(4).BackColor = QBColor(7)
        Case 5
        importe(5).BackColor = QBColor(7)
        trabajo(5).BackColor = QBColor(7)
        Case 6
        importe(6).BackColor = QBColor(7)
        trabajo(6).BackColor = QBColor(7)
        Case 7
        importe(7).BackColor = QBColor(7)
        trabajo(7).BackColor = QBColor(7)
        Case 8
        importe(8).BackColor = QBColor(7)
        trabajo(8).BackColor = QBColor(7)
        Case 9
        importe(9).BackColor = QBColor(7)
        trabajo(9).BackColor = QBColor(7)
        End Select
End Sub

Private Sub importe_LostFocus(Index As Integer)
Select Case Index
       Case 0
        importe(0).BackColor = QBColor(15)
        trabajo(0).BackColor = QBColor(15)
        Case 1
        importe(1).BackColor = QBColor(15)
        trabajo(1).BackColor = QBColor(15)
        Case 2
        importe(2).BackColor = QBColor(15)
        trabajo(2).BackColor = QBColor(15)
        Case 3
        importe(3).BackColor = QBColor(15)
        trabajo(3).BackColor = QBColor(15)
        Case 4
        importe(4).BackColor = QBColor(15)
        trabajo(4).BackColor = QBColor(15)
        Case 5
        importe(5).BackColor = QBColor(15)
        trabajo(5).BackColor = QBColor(15)
        Case 6
        importe(6).BackColor = QBColor(15)
        trabajo(6).BackColor = QBColor(15)
        Case 7
        importe(7).BackColor = QBColor(15)
        trabajo(7).BackColor = QBColor(15)
        Case 8
        importe(8).BackColor = QBColor(15)
        trabajo(8).BackColor = QBColor(15)
        Case 9
        importe(9).BackColor = QBColor(15)
        trabajo(9).BackColor = QBColor(15)
        End Select
End Sub

Private Sub importe_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("|") Then
    KeyAscii = 0
Else
    If KeyAscii <> 8 And blibres() <= 0 Then
        KeyAscii = 0
    End If
End If
End Sub


Private Sub trabajo_GotFocus(Index As Integer)
Select Case Index
        Case 0
        trabajo(0).BackColor = QBColor(7)
        importe(0).BackColor = QBColor(7)
        Case 1
        trabajo(1).BackColor = QBColor(7)
        importe(1).BackColor = QBColor(7)
        Case 2
        trabajo(2).BackColor = QBColor(7)
        importe(2).BackColor = QBColor(7)
        Case 3
        trabajo(3).BackColor = QBColor(7)
        importe(3).BackColor = QBColor(7)
        Case 4
        trabajo(4).BackColor = QBColor(7)
        importe(4).BackColor = QBColor(7)
        Case 5
        trabajo(5).BackColor = QBColor(7)
        importe(5).BackColor = QBColor(7)
        Case 6
        trabajo(6).BackColor = QBColor(7)
        importe(6).BackColor = QBColor(7)
        Case 7
        trabajo(7).BackColor = QBColor(7)
        importe(7).BackColor = QBColor(7)
        Case 8
        trabajo(8).BackColor = QBColor(7)
        importe(8).BackColor = QBColor(7)
        Case 9
        trabajo(9).BackColor = QBColor(7)
        importe(9).BackColor = QBColor(7)
        End Select
End Sub

Private Sub trabajo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("|") Then
    KeyAscii = 0
Else
    If KeyAscii <> 8 And blibres() = 10 Then
        KeyAscii = 0
    End If
End If
End Sub

Function blibres()
Dim i As Byte
Dim total As Long
Const delims = 19
Const tamtotal = 1024

total = 0
For i = 0 To 9
    total = total + (Len(trabajo(i)) + Len(importe(i)))
Next i
total = total + delims
blibres = tamtotal - total
Label22.Caption = str(blibres)
End Function


Private Sub trabajo_LostFocus(Index As Integer)
Select Case Index
       Case 0
        trabajo(0).BackColor = QBColor(15)
        importe(0).BackColor = QBColor(15)
        Case 1
        trabajo(1).BackColor = QBColor(15)
        importe(1).BackColor = QBColor(15)
        Case 2
        trabajo(2).BackColor = QBColor(15)
        importe(2).BackColor = QBColor(15)
        Case 3
        trabajo(3).BackColor = QBColor(15)
        importe(3).BackColor = QBColor(15)
        Case 4
        trabajo(4).BackColor = QBColor(15)
        importe(4).BackColor = QBColor(15)
        Case 5
        trabajo(5).BackColor = QBColor(15)
        importe(5).BackColor = QBColor(15)
        Case 6
        trabajo(6).BackColor = QBColor(15)
        importe(6).BackColor = QBColor(15)
        Case 7
        trabajo(7).BackColor = QBColor(15)
        importe(7).BackColor = QBColor(15)
        Case 8
        trabajo(8).BackColor = QBColor(15)
        importe(8).BackColor = QBColor(15)
        Case 9
        trabajo(9).BackColor = QBColor(15)
        importe(9).BackColor = QBColor(15)
        End Select
End Sub
