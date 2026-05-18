VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   6120
         Top             =   2640
      End
      Begin VB.Image imgLogo 
         Height          =   3945
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Copyright DarkSoft Dev 1992-2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label lblWarning 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Advertencia : Este prgorama esta protegido por las leyes de derecho de autor y otros tratados internacionales."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2280
         TabIndex        =   1
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Para Plataformas Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   2520
         Width           =   2685
      End
      Begin VB.Label lblProductName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "FluWork 10 D2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
