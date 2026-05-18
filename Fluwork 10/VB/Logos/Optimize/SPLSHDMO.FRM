VERSION 5.00
Begin VB.Form frmSplashDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usar una pantalla inicial"
   ClientHeight    =   1530
   ClientLeft      =   1515
   ClientTop       =   1470
   ClientWidth     =   5910
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   102
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   394
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAppStartup 
      Caption         =   "&Sin pantalla inicial"
      Height          =   500
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   1635
   End
   Begin VB.CommandButton cmdAppStartup 
      Caption         =   "&Con pantalla inicial"
      Height          =   500
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usar una pantalla inicial proporciona al usuario información mientras se carga su aplicación."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   30
      TabIndex        =   2
      Top             =   45
      Width           =   5790
   End
End
Attribute VB_Name = "frmSplashDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAppStartup_Click(Index As Integer)
  Dim dStartTime As Double
  
  dStartTime = Timer
  
    Select Case Index
      Case 0  'Con ventana de presentación
        ShowSplash
      Case 1  'Sin ventana de presentación
        NoSplash
    End Select
    
    MsgBox "Tiempo de inicio = " & Format$(Timer - dStartTime, "##.##") & " segs.", _
                                            vbInformation, _
                                            "Timpo de inicio de la aplicación"
  End Sub

Private Sub NoSplash()
            
  ' Ahora carga un montón de formularios.
  Dim foo1 As New frmImages
  foo1.Caption = "Foo1"
  foo1.Left = 0
  foo1.Top = 0
          
  Dim foo2 As New frmImages
  foo2.Caption = "Foo2"
  foo2.Left = 200
  foo2.Top = 100
          
  Dim foo3 As New frmImages
  foo3.Caption = "Foo3"
  foo3.Left = 300
  foo3.Top = 150
          
  Dim foo4 As New frmImages
  foo4.Caption = "Foo4"
  foo4.Left = 400
  foo4.Top = 200

End Sub


Private Sub ShowSplash()
  Dim success%
  Dim iStatusBarWidth As Integer
  
  On Error GoTo SplashLoadErr
  
  iStatusBarWidth = 4575
            
  Screen.MousePointer = vbHourglass
  Load Splash
  Splash.Show
  DoEvents
  ' Establece que la pantalla de impacto permanezca en primer plano.
  success% = SetWindowPos(Splash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

  ' Ahora carga un montón de formularios.
  Dim foo1 As New frmImages
  'Splash.Refresh
  foo1.Caption = "Foo1"
  foo1.Move 0, 0
  DoEvents
  Splash.rctStatusBar.Width = iStatusBarWidth * 0.25
          
  Dim foo2 As New frmImages
  'Splash.Refresh
  foo2.Caption = "Foo2"
  foo2.Move 0, 0
  DoEvents
  Splash.rctStatusBar.Width = iStatusBarWidth * 0.5

  Dim foo3 As New frmImages
  'Splash.Refresh
  foo3.Caption = "Foo3"
  foo3.Move 0, 0
  DoEvents
  Splash.rctStatusBar.Width = iStatusBarWidth * 0.75
          
  Dim foo4 As New frmImages
  'Splash.Refresh
  foo4.Caption = "Foo4"
  foo4.Move 0, 0
  DoEvents
  Splash.rctStatusBar.Width = iStatusBarWidth

  ' Desactiva el indicador de primer plano de la ventana.
  success% = SetWindowPos(Splash.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  Unload Splash
  Screen.MousePointer = vbDefault
  Exit Sub

SplashLoadErr:
  success% = SetWindowPos(Splash.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  Unload Splash
  Screen.MousePointer = vbDefault
  MsgBox Error$ & " - " & Str$(Err), vbExclamation, "Error al cargar la aplicación"
  Exit Sub

End Sub
Private Sub Form_Load()
  Me.Left = frmExplore.Width + 400
  Me.Top = (Screen.Height - Me.Height) * 0.9
End Sub


