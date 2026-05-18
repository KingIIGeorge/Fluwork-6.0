VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Peonza roja"
   ClientHeight    =   2325
   ClientLeft      =   1365
   ClientTop       =   1755
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Redtop.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   239
   Begin PicClip.PictureClip PictureClip2 
      Left            =   120
      Top             =   1920
      _ExtentX        =   10398
      _ExtentY        =   4815
      _Version        =   327680
      Rows            =   3
      Cols            =   6
      Picture         =   "Redtop.frx":030A
   End
   Begin VB.CommandButton onoff 
      Caption         =   "Girar"
      Height          =   390
      Left            =   1725
      TabIndex        =   1
      Top             =   1095
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Información"
      Height          =   345
      Left            =   1725
      TabIndex        =   2
      Top             =   705
      Width           =   1245
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00408080&
      ForeColor       =   &H00000000&
      Height          =   1245
      Left            =   225
      ScaleHeight     =   79
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   86
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1350
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   60
         ScaleHeight     =   46
         ScaleMode       =   3  'Píxel
         ScaleWidth      =   53
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   60
         Width           =   825
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2235
      Top             =   120
   End
   Begin VB.PictureBox PictureClip1 
      Height          =   480
      Left            =   -3135
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   17
      Left            =   3720
      Picture         =   "Redtop.frx":91CC
      Top             =   4425
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   16
      Left            =   3720
      Picture         =   "Redtop.frx":94D6
      Top             =   3885
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   15
      Left            =   3660
      Picture         =   "Redtop.frx":97E0
      Top             =   3345
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   14
      Left            =   3660
      Picture         =   "Redtop.frx":9AEA
      Top             =   2805
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   13
      Left            =   3660
      Picture         =   "Redtop.frx":9DF4
      Top             =   2265
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   3660
      Picture         =   "Redtop.frx":A0FE
      Top             =   1725
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   3660
      Picture         =   "Redtop.frx":A408
      Top             =   1185
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   3660
      Picture         =   "Redtop.frx":A712
      Top             =   645
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   3660
      Picture         =   "Redtop.frx":AA1C
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   3120
      Picture         =   "Redtop.frx":AD26
      Top             =   4425
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   3120
      Picture         =   "Redtop.frx":B030
      Top             =   3885
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   3120
      Picture         =   "Redtop.frx":B33A
      Top             =   3345
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   3120
      Picture         =   "Redtop.frx":B644
      Top             =   2805
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   3120
      Picture         =   "Redtop.frx":B94E
      Top             =   2265
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   3120
      Picture         =   "Redtop.frx":BC58
      Top             =   1725
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   3120
      Picture         =   "Redtop.frx":BF62
      Top             =   1185
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   3120
      Picture         =   "Redtop.frx":C26C
      Top             =   645
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   3120
      Picture         =   "Redtop.frx":C576
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim y As Integer
Dim x As Integer
Dim delay As Integer
Dim toggle As Integer
Dim flap As Integer

Private Sub Command1_Click()
    infoform.Show   ' Carga y muestra el formulario de información.
End Sub

Private Sub Form_Load()
    ' Usa el tercer marco de la secuencia como marco de inicio.
    Picture1.Picture = PictureClip2.GraphicCell(2)
    y = 2
    ' Centra la imagen más pequeña en la imagen más grande.
    Picture1.Left = (Picture6.ScaleWidth - Picture1.Width) / 2
    Picture1.Top = (Picture6.ScaleHeight - Picture1.Height) / 2
End Sub

Private Sub onoff_Click()
    If toggle = 0 Then
        onoff.Caption = "Parar"
        toggle = 1
    Else
        onoff.Caption = "Girar"
        toggle = 0
    End If
End Sub

Private Sub runtop()
    ' Avanza la animación una imagen.
    y = y + 1: If y = 18 Then y = 0
    
    Picture1.Picture = PictureClip2.GraphicCell(y)
    
    ' Animación de icono:  sólo se verá cuando el formulario esté
    ' minimizado. En la matriz Picture3 se han cargado archivos de
    ' icono (.ICO), y no .BMP. Esto permite usar la función de
    ' máscara del archivo de icono, dejando que el fondo que hay
    ' tras el icono se muestre a través de él.
    
    Form1.Icon = Image1(y).Picture

End Sub

Private Sub Timer1_Timer()
    If toggle = 1 Then runtop
End Sub

