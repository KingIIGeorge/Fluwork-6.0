VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Optimizar Visual Basic 5.0"
   ClientHeight    =   2205
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "ABOUT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2205
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton ExitButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   135
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "ABOUT.frx":030A
      Top             =   195
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   1020
      X2              =   5100
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1020
      X2              =   5100
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicación de ejemplo de optimización de código de Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2835
   End
   Begin VB.Label lblCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " © 1995 - Microsoft Corporation"
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   900
      TabIndex        =   1
      Top             =   1395
      Width           =   3315
   End
   Begin VB.Label lblAppInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   900
      TabIndex        =   2
      Top             =   195
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ExitButton_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    CenterMe Me
    DoEvents
    
    lblAppInfo = "Optimizar Visual Basic 5.0" & _
                      vbCrLf & _
                      "Versión " & App.Major & "." & App.Minor & _
                      " - Build " & LastAppUpdate

End Sub

Private Sub Picture1_DblClick()
    
End Sub

