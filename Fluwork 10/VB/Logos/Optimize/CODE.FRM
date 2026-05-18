VERSION 5.00
Begin VB.Form Code 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Optimización de código"
   ClientHeight    =   5220
   ClientLeft      =   225
   ClientTop       =   690
   ClientWidth     =   7020
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5220
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Tag             =   "Real"
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtLoopCount 
      Height          =   320
      Left            =   5535
      TabIndex        =   8
      Text            =   "txtLoopCount"
      Top             =   3765
      Width           =   1300
   End
   Begin VB.CommandButton cmdExplicit 
      Caption         =   "Propiedad explícita"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdVariable 
      Caption         =   "Prueba de variables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdVariant 
      Caption         =   "Contador de tipo Variant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   4
      Top             =   4190
      Width           =   1935
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "Contador de tipo Integer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtValue 
      Height          =   320
      Left            =   5535
      TabIndex        =   7
      Text            =   "txtValue"
      Top             =   1800
      Width           =   1300
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "Valor del control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   2780
      Width           =   1935
   End
   Begin VB.TextBox txtCounter 
      Height          =   320
      Left            =   5535
      TabIndex        =   6
      Text            =   "txtCounter"
      Top             =   60
      Width           =   1300
   End
   Begin VB.CommandButton cmdProperty 
      Caption         =   "Prueba de propiedad"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   60
      X2              =   6960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   60
      X2              =   6960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   60
      X2              =   6960
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Variables Integer contra Variant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   3750
      Width           =   3675
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Propiedad explícita contra valor del control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   5550
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Acceso a variables y a propiedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblVariant 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00,00 segs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5535
      TabIndex        =   18
      Top             =   4290
      Width           =   1380
   End
   Begin VB.Label lblInteger 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00,00 segs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5535
      TabIndex        =   17
      Top             =   4785
      Width           =   1380
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usando la propiedad Text:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2145
      TabIndex        =   16
      Top             =   2415
      Width           =   3960
   End
   Begin VB.Label lblExplicit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00,00 segs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5535
      TabIndex        =   15
      Top             =   2385
      Width           =   1380
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usando el valor del control:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2145
      TabIndex        =   14
      Top             =   2895
      Width           =   2700
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00,00 segs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5535
      TabIndex        =   13
      Top             =   2880
      Width           =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   60
      X2              =   6960
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblVariable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00,00 segs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5535
      TabIndex        =   12
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usando una variable en el bucle:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2130
      TabIndex        =   11
      Top             =   1095
      Width           =   3105
   End
   Begin VB.Label lblProperty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00,00 segs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5535
      TabIndex        =   10
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usando la propiedad Text en el bucle:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2130
      TabIndex        =   9
      Top             =   615
      Width           =   3345
   End
End
Attribute VB_Name = "Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExplicit_Click()
'Utiliza la referencia explícita a la propiedad (txtValue.Text)
'La subrutina cmdValue_Click es algo más rápida
    Limit% = Val(txtValue)
    Screen.MousePointer = 11
    start = Timer
    For x% = 1 To Limit%
        foo$ = txtValue.Text
    Next x%
    Finish = Timer
    Screen.MousePointer = 0
    lblExplicit = Format$(Finish - start, "##.##") & " segs."
End Sub

Private Sub cmdInteger_Click()
'Utiliza un entero como contador del bucle,
'es más rápido que la subrutina cmdVariant_Click
Dim intCounter1 As Integer, intCounter2 As Integer, intCounter3 As Integer
    Screen.MousePointer = 11
    start = Timer
    For intCounter1 = 1 To 100
        For intCounter2 = 1 To intCounter1
            For intCounter3 = 1 To intCounter2
                Var = intCounter3 * 100 + intCounter2 * 10 + intCounter1
            Next intCounter3
        Next intCounter2
    Next intCounter1
    Finish = Timer
    Screen.MousePointer = 0
    lblInteger = Format$(Finish - start, "##.##") & " segs."
End Sub

Private Sub cmdProperty_Click()
'En cada iteración, hace referencia a la propiedad Text
'La subrutina cmdVariable_Click es más rápida
    Limit% = Val(txtCounter.Text)
    Screen.MousePointer = 11
    
    start = Timer
    
    For x% = 1 To Limit%
        foo$ = Left(txtCounter.Text, Len(txtCounter.Text) \ 2)
    Next x%
    
    Finish = Timer
    
    Screen.MousePointer = 0
    lblProperty = Format$(Finish - start, "##.##") & " segs."
End Sub

Private Sub cmdValue_Click()
'Utiliza el valor del control al omitir la propiedad explícita Text,
'es algo más rápido que la subrutina cmdExplicit_Click
    Limit% = Val(txtValue)
    Screen.MousePointer = 11
    start = Timer
    For x% = 1 To Limit%
        foo$ = txtValue
    Next x%
    Finish = Timer
    Screen.MousePointer = 0
    lblValue = Format$(Finish - start, "##.##") & " segs."
End Sub

Private Sub cmdVariable_Click()
'Utiliza una variable (Temp$) en lugar del acceso a la propiedad,
'es más rápido que la subrutina cmdProperty_Click
    Limit% = Val(txtCounter.Text)
    Temp$ = txtCounter.Text
    Screen.MousePointer = 11
    
    start = Timer
    
    For x% = 1 To Limit%
        foo$ = Left(Temp$, Len(Temp$) \ 2)
    Next x%
    
    Finish = Timer
    
    Screen.MousePointer = 0
    lblVariable = Format$(Finish - start, "##.##") & " segs."
End Sub
Private Sub cmdVariant_Click()
'Utiliza una variable de tipo Variant como contador del bucle.
'La subrutina cmdInteger_Click es más rápida puesto que utiliza un entero como contador del bucle
    
    Screen.MousePointer = 11
    start = Timer
    
    For varCounter1 = 1 To txtLoopCount.Text
        For varCounter2 = 1 To varCounter1
            For varCounter3 = 1 To varCounter2
                Var = varCounter3 * 100 + varCounter2 * 10 + varCounter1
            Next varCounter3
        Next varCounter2
    Next varCounter1
    
    Finish = Timer
    Screen.MousePointer = 0
    lblVariant = Format$(Finish - start, "##.##") & " segs."
End Sub

Private Sub Form_Load()
    PosForm Me
    
    ' Establece los valores predeterminados de los contadores
    txtCounter = 1000
    txtValue = 1000
    txtLoopCount = 100

End Sub

