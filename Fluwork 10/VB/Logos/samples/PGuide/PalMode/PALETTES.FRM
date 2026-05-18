VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demostración de PaletteMode"
   ClientHeight    =   5415
   ClientLeft      =   1860
   ClientTop       =   1530
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6465
   Begin VB.OptionButton Option1 
      Caption         =   "Personalizado (DIB)"
      Height          =   255
      Index           =   3
      Left            =   3300
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   4380
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3795
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   1020
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Personalizado (escala de grises)"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2715
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Orden Z"
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   1
      Top             =   240
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Semitonos"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   5160
      Top             =   3180
   End
   Begin VB.Label Label1 
      Caption         =   "Nota: Para un efecto real, debe ejecutar esta demostración en 256 colores."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5100
      Width           =   6135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1290
      Left            =   4680
      Picture         =   "Palettes.frx":0000
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Aplicación de ejemplo para demostrar las propiedades PaletteMode          ///
'// y Palette de los formularios y los controles.                             ///
'// Versión 1.0 9/16/1996                                                     ///
'//                                                                           ///
'// Nota: este ejemplo tiene que ser ejecutado en el modo de 256 colores      ///
'// Palette y PaletteMode no tienen ningún efecto en otros                    ///
'// modos de color.

Option Explicit
Dim objPic As Picture    ' Objeto imagen de la opción 4.

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Carga un DIB de 256 colores en el objeto imagen
    Set objPic = LoadPicture(App.Path & "\PASTEL.DIB")
End Sub

Private Sub Option1_Click(Index As Integer)
    Timer1.Enabled = False
    Select Case Index
        Case 0
            ' Utiliza el modo Halftone (predeterminado).
            Form1.PaletteMode = vbPaletteModeHalftone
        Case 1
            ' Utiliza la paleta de la imagen cargada -
            ' se producirá un cambio de color (modo ZOrder).
            Form1.PaletteMode = vbPaletteModeUseZOrder
        Case 2
            'Set Form1.Palette = Nothing
            ' Asigna la paleta de Image1 al formulario.
            Form1.Palette = Image1.Picture
            ' Utiliza el modo Custom.
            Form1.PaletteMode = vbPaletteModeCustom
        Case 3
            'Set Form1.Palette = Nothing
            ' Asigna la paleta del objeto imagen al formulario.
            Form1.Palette = objPic
            ' Utiliza el modo Custom.
            Form1.PaletteMode = vbPaletteModeCustom
    End Select
    Picture1.Refresh
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static intC As Integer
    
    ' Cambia entre tres imágenes diferentes.
    If intC < 1 Then
        Picture1 = LoadPicture(App.Path & "\BANNER.GIF")
        intC = 1
    ElseIf intC = 1 Then
        Picture1 = LoadPicture(App.Path & "\CLOUDS.BMP")
        intC = 2
    Else
        Picture1 = LoadPicture(App.Path & "\FOREST.JPG")
        intC = 0
    End If
End Sub

