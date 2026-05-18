VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ejemplos de controles"
   ClientHeight    =   2910
   ClientLeft      =   210
   ClientTop       =   1935
   ClientWidth     =   4830
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
   ScaleHeight     =   2910
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdImages 
      Caption         =   "&Imágenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   1755
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "Botones de &opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1755
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Casilla de verificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1755
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Cuadro de te&xto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1755
   End
   Begin VB.CommandButton cmdWordWrap 
      Caption         =   "&Word Wrap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1755
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "&Prueba de botones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1755
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   1755
   End
   Begin VB.Menu mnuFileMain 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuOptionsMain 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuButtons 
         Caption         =   "&Prueba de botones"
      End
      Begin VB.Menu mnuWordWrap 
         Caption         =   "&WordWrap"
      End
      Begin VB.Menu mnuText 
         Caption         =   "Propiedades del cuadro de te&xto"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "&Casilla de verificación"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Botones de &opción"
      End
      Begin VB.Menu mnuImages 
         Caption         =   "&Imágenes"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdButtons_Click()
    ' invoca un evento Click en el menú
    mnuButtons_Click
End Sub

Private Sub cmdCheck_Click()
    ' invoca un evento Click en el menú
    mnuCheck_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me       ' descarga el formulario
    End             ' termina la aplicación
End Sub

Private Sub cmdImages_Click()
    ' invoca un evento Click en el menú
    mnuImages_Click
End Sub

Private Sub cmdOption_Click()
    ' invoca un evento Click en el menú
    mnuOption_Click
End Sub

Private Sub cmdText_Click()
    ' invoca un evento Click en el menú
    mnuText_Click
End Sub

Private Sub cmdWordWrap_Click()
    ' invoca un evento Click en el menú
    mnuWordWrap_Click
End Sub

Private Sub Form_Load()
    frmMain.Height = 3600
    frmMain.Width = 4965
End Sub

Private Sub mnuButtons_Click()
    ' presenta el formulario
    frmButton.Show
End Sub

Private Sub mnuCheck_Click()
    ' presenta el formulario
    frmCheck.Show
End Sub

Private Sub mnuFileExit_Click()
    ' invoca un evento Click en el botón de comando
    cmdExit_Click
End Sub

Private Sub mnuImages_Click()
    ' presenta el formulario
    frmImages.Show
End Sub

Private Sub mnuOption_Click()
    ' presenta el formulario
    frmOptions.Show
End Sub

Private Sub mnuText_Click()
    ' presenta el formulario
    frmText.Show
End Sub

Private Sub mnuWordWrap_Click()
    ' presenta el formulario
    frmWordWrap.Show
End Sub

