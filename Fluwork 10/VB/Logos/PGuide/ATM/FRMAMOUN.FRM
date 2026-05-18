VERSION 5.00
Begin VB.Form frmAmountWithdrawn 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3045
   ClientLeft      =   3375
   ClientTop       =   4065
   ClientWidth     =   4155
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4155
   Begin VB.Frame fraMoney 
      Caption         =   "fraMoney"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
      Begin VB.TextBox txtConvertedAmt 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtUSDollarsAmt 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCurrency 
         Caption         =   "lblCurrency"
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblUSDollars 
         Caption         =   "US Dollars"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOKEnd 
      Caption         =   "cmdOKEnd"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgFlag 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblAmountWith 
      Caption         =   "lblAmountWith"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmAmountWithdrawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOKEnd_Click()
    ' Para reproducir un sonido, es necesario una tarjeta de sonido con un controlador Microsoft ADPCM Audio CODEC.
    ' Este controlador se incluye en Windows 95 y Windows NT 3.51. Para Windows 3.1x,
    ' el controlador CODEC viene incluido en Video for Windows 1.1 o Windows Sound System 2.0.
    BeginPlaySound i
    MsgBox LoadResString(8 + i)   ' Presenta un mensaje de agradecimiento.
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ConversionValue As Currency
    ConversionValue = ConversionTable((i - 16) \ 32)
    imgFlag.Picture = LoadResPicture(0 + i, vbResBitmap)
    lblAmountWith.Caption = LoadResString(9 + i) & " " & _
        IIf(frmInput.optChecking.Value, LoadResString(3 + i), LoadResString(4 + i))
    lblCurrency.Caption = LoadResString(10 + i)
    cmdOKEnd.Caption = LoadResString(6 + i)
    fraMoney.Caption = "1 USD ($) = " & CStr(ConversionValue) & " " & lblCurrency.Caption
    txtUSDollarsAmt.Text = Val(frmInput.txtUSDollarsAmt.Text)
    txtConvertedAmt.Text = ConversionValue * CCur(txtUSDollarsAmt.Text)
    SetCursor cmdOKEnd
    
    ' En DBCS, el nombre de la fuente tiene que ser establecido por VB al cargar
    ' el proyecto. Para hacer que este proyecto funcione en todas partes,
    ' tenemos que establecer las propiedades de la fuente desde el programa.
    lblAmountWith.FontBold = True
End Sub
