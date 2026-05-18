VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Establecer contraseña"
   ClientHeight    =   1995
   ClientLeft      =   3645
   ClientTop       =   3540
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016136
   Icon            =   "NEWPW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVerify 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtNewPW 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtOldPW 
      Height          =   285
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Comprobar:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nueva contraseña:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1590
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña anterior:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Establecer contraseña"
Const BUTTON1 = "&Aceptar"
Const BUTTON2 = "&Cancelar"
Const Label1 = "C&ontraseña anterior:"
Const Label2 = "&Nueva contraseña:"
Const LABEL3 = "Co&mprobar:"
Const MSG1 = "Error al comprobar la contraseña, vuelva a intentarlo."
'>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo OKErr
  
  Dim sTmp As String
  Dim usr As User
  
  If txtNewPW.Text <> txtVerify.Text Then
    Beep
    MsgBox MSG1, 48
    txtVerify.SetFocus
    Exit Sub
  End If
  
  gwsMainWS.Users(gwsMainWS.UserName).NewPassword txtOldPW.Text, txtNewPW.Text
  
  Unload Me
  Exit Sub
  
OKErr:
  ShowError
End Sub

Private Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
  lblLabels(2).Caption = LABEL3
End Sub
