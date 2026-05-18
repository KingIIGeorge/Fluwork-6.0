VERSION 5.00
Begin VB.Form frmExpName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar"
   ClientHeight    =   1335
   ClientLeft      =   2775
   ClientTop       =   2670
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1800
   HelpContextID   =   2016127
   Icon            =   "EXPNAME.frx":0000
   Left            =   2715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Top             =   2265
   Width           =   5070
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtTable 
      Height          =   285
      Left            =   160
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmExpName"
Attribute VB_Base = "0{529A4483-C9E1-11CF-9ED2-00AA00574745}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Customizable = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Exportar"
Const BUTTON1 = "&Aceptar"
Const BUTTON2 = "&Cancelar"
'>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub cmdCancel_Click()
  gExpTable = vbNullString
  Unload Me
End Sub

Private Sub cmdOK_Click()
  gExpTable = txtTable.Text
  Unload Me
End Sub

Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
End Sub
