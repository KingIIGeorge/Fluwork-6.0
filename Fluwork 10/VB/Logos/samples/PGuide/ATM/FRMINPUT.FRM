VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3645
   ClientLeft      =   1800
   ClientTop       =   3030
   ClientWidth     =   5220
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5220
   Begin VB.Frame fraAccount 
      Caption         =   "fraAccount"
      Height          =   1215
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   3975
      Begin VB.OptionButton optChecking 
         Caption         =   "optChecking"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optSavings 
         Caption         =   "optSavings"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "cmdOK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtUSDollarsAmt 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtPin 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblUSDollars 
      Caption         =   "US Dollars"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   2700
      Width           =   1575
   End
   Begin VB.Image imgFlag 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblAmount 
      Caption         =   "lblAmount"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label lblPINCode 
      Caption         =   "lblPINCode"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_Base = "0{DC791295-F602-11CF-8A54-0020AF939EAB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Esta aplicación de ejemplo muestra cómo usar archivos de recursos para localizar
' una aplicación Visual Basic. No usa controladores de errores o de validación de datos.

' Esta aplicación usa la función LoadResPicture para cargar una imagen, LoadResString
' para cargar una cadena, y LoadResData para leer un archivo .WAV del archivo de recursos,
' dependiendo del botón de comando que elija el usuario en el evento frmOpen.
Option Explicit

Sub Form_Load()
    imgFlag = LoadResPicture(I, vbResBitmap)
    Caption = LoadResString(I)
    lblPINCode = LoadResString(1 + I)
    fraAccount = LoadResString(2 + I)
    optChecking.Caption = LoadResString(3 + I)
    optSavings.Caption = LoadResString(4 + I)
    lblAmount = LoadResString(5 + I)
    cmdOK.Caption = LoadResString(6 + I)
    SetCursor cmdOK
End Sub

Sub cmdOK_click()
    MsgBox LoadResString(7 + I)   ' Muestra un mensaje de proceso.
    frmAmountWithdrawn.Show vbModal
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndPlaySound
End Sub

