VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1170
      TabIndex        =   1
      Top             =   135
      Width           =   2445
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2445
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establece la variable global a false
    'para indicar un fallo en el inicio de sesión
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'comprueba la contraseña correcta
    If txtPassword = "password" Then
        'sitúe código aquí para indicar que
        'se conoce la contraseña
        'lo más fácil es establecer una variable global
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Contraseña no válida, vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

