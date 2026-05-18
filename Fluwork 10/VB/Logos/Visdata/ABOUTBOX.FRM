VERSION 5.00
Begin VB.Form frmAboutBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Visual Data Manager"
   ClientHeight    =   2280
   ClientLeft      =   4395
   ClientTop       =   2925
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   2745
   Icon            =   "ABOUTBOX.frx":0000
   Left            =   4335
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Top             =   2520
   Width           =   5175
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   255
      Picture         =   "ABOUTBOX.frx":000C
      ScaleHeight     =   525
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Copyright Microsoft Corp. 1996-97"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Versión 5.0"
      Height          =   210
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Visual Data Manager"
      Height          =   210
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Microsoft Visual Basic"
      Height          =   270
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_Base = "0{529A4441-C9E1-11CF-9ED2-00AA00574745}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Customizable = False
Option Explicit

'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Acerca de Visual Data Manager"
Const LABEL0 = "Microsoft Visual Basic"
Const Label1 = "Visual Data Manager"
Const Label2 = "Versión 5.0"
Const LABEL3 = "Copyright Microsoft Corp. 1996-97"
Const BUTTON1 = "&Aceptar"
'>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub Form_KeyPress(KeyAscii As Integer)
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = FORMCAPTION
  lblLabels(0).Caption = LABEL0
  lblLabels(1).Caption = Label1
  lblLabels(2).Caption = Label2
  lblLabels(3).Caption = LABEL3
  cmdOK.Caption = BUTTON1
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
  Unload Me
End Sub
