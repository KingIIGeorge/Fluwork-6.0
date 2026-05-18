VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   1980
   ClientLeft      =   2550
   ClientTop       =   3330
   ClientWidth     =   2820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016133
   Icon            =   "Loginfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1440
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtLoginName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña: "
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de inicio de sesión: "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Inicio de sesión"
Const BUTTON1 = "&Aceptar"
Const BUTTON2 = "&Cancelar"
Const Label1 = "Nombre de inicio de sesión: "
Const Label2 = "Contraseña:"
'>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub cmdCancel_Click()
  On Error Resume Next
  'si el objeto Workspace global no ha sido establecido, debemos estar 
  'iniciando la aplicación, de modo que deberíamos terminar
  If gwsMainWS Is Nothing Then
    End
  Else
    'descarga porque el Workspace ya está establecido
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
End Sub

Private Sub cmdOK_Click()
  On Error GoTo OKErr
  
  Dim wsp As Workspace
  Dim sTmp As String

  If Not gwsMainWS Is Nothing Then
    If UCase(txtLoginName.Text) = UCase(gwsMainWS.UserName) Then
      'igual que el nombre de sesión actual
      Unload Me
      Exit Sub
    End If
  End If
  
  'establece el nuevo nombre de inicio de sesión
  DBEngine.DefaultUser = txtLoginName.Text
  DBEngine.DefaultPassword = txtPassword.Text
  
  Set wsp = DBEngine.CreateWorkspace("MainWS", txtLoginName.Text, txtPassword.Text)
  'ha sido correcto, luego establecemos gswMainWS
  Set gwsMainWS = wsp
  Unload Me
  
  Exit Sub

OKErr:
  MsgBox Error
  txtLoginName.SetFocus
  Exit Sub        'ofrecemos otra oportunidad

End Sub

Private Sub txtLoginName_GotFocus()
  txtLoginName.SelStart = 0
  txtLoginName.SelLength = Len(txtLoginName.Text)
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword.Text)
End Sub
