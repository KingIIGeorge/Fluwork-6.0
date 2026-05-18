VERSION 5.00
Begin VB.Form frmWordWrap 
   Caption         =   "Propiedades WordWrap y Autosize"
   ClientHeight    =   4320
   ClientLeft      =   1110
   ClientTop       =   1665
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdCycle 
      Caption         =   "&Ciclo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox chkAutoSize 
      Caption         =   "&AutoSize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkWordWrap 
      Caption         =   "&WordWrap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&errar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccione las casillas de verificación AutoSize y WordWrap para ver su efecto en la etiqueta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Una demostración de las propiedades AutoSize y WordWrap."
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmWordWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Display()
    ' esta subrutina modifica las características de
    ' presentación del control Label
    
    ' Restablece el programa.
    lblDisplay.AutoSize = False
    lblDisplay.WordWrap = False
    lblDisplay.Width = 1695
    lblDisplay.Height = 225
    ' Comprueba los valores de WordWrap y Autosize.
    If chkWordWrap.Value = 1 Then
        lblDisplay.WordWrap = True
    End If
    If chkAutoSize.Value = 1 Then
        lblDisplay.AutoSize = True
    End If
End Sub

Private Sub chkAutoSize_Click()
    ' llama a la subrutina Display
    Call Display
End Sub


Private Sub chkWordWrap_Click()
    ' llama a la subrutina Display
    Call Display
End Sub


Private Sub cmdClose_Click()
   Unload Me    ' Descarga este formulario.
End Sub

Private Sub cmdCycle_Click()
    ' recorre las cuatro combinaciones posibles
    
    ' ninguna de las casillas de verificación está seleccionada
    If chkAutoSize.Value = 0 And _
     chkWordWrap.Value = 0 Then
        ' selecciona la casilla de verificación AutoSize
        chkAutoSize.Value = 1
    ' ambas casillas de verificación están seleccionadas
    ElseIf chkAutoSize.Value = 1 And _
     chkWordWrap.Value = 1 Then
        ' borra la casilla de verificación AutoSize y
        ' selecciona la casilla de verificación WordWrap
        chkAutoSize.Value = 0
        chkWordWrap.Value = 1
    ' sólo la casilla de verificación WordWrap está seleccionada
    ElseIf chkAutoSize.Value = 0 And _
     chkWordWrap.Value = 1 Then
        ' borra las dos casillas de verificación
        chkAutoSize.Value = 0
        chkWordWrap.Value = 0
    ' sólo la casilla de verificación AutoSize está seleccionada
    Else
        ' selecciona la casilla de verificación WordWrap. La
        ' casilla de verificación AutoSize ya está seleccionada
        chkWordWrap.Value = 1
    End If
End Sub


