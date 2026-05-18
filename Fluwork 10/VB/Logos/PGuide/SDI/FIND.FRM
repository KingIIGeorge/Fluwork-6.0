VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   1395
   ClientLeft      =   1500
   ClientTop       =   1875
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5850
   Begin VB.Frame Frame1 
      Caption         =   "Dirección"
      Height          =   612
      Left            =   2400
      TabIndex        =   3
      Top             =   660
      Width           =   2052
      Begin VB.OptionButton optDirection 
         Caption         =   "&Abajo"
         Height          =   252
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "Buscar hasta el final del documento"
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "A&rriba"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Buscar hasta el principio del documento"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "&Coincidencia exacta"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Coincidir MAY/min"
      Top             =   840
      Width           =   2115
   End
   Begin VB.TextBox txtFind 
      Height          =   500
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Texto para buscar"
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   372
      Left            =   4620
      TabIndex        =   7
      ToolTipText     =   "Volver al Bloc de notas"
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   372
      Left            =   4620
      TabIndex        =   6
      ToolTipText     =   "Iniciar búsqueda"
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "B&uscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Cuadro de diálogo Buscar para búsquedas de texto.               ***
'*** Creado para la aplicación de ejemplo Bloc de notas SDI          ***
'*** Utiliza las variables públicas gFindCase (alterna la distinción ***
'*** entre mayúsculas y minúsculas); gFindString (texto que se       ***
'*** busca); gFindDirection (alterna la dirección de la              ***
'*** búsqueda); gFirstTime (alterna el inicio desde el principio     ***
'*** del texto)                                                      ***
'***********************************************************************

Option Explicit

Private Sub chkCase_Click()
    ' Asigna un valor a la variable pública.
    gFindCase = chkCase.Value
End Sub

Private Sub cmdCancel_Click()
    ' Guarda los valores en las variables públicas.
    gFindString = txtFind.Text
    gFindCase = chkCase.Value
    ' Descarga el cuadro de diálogo de búsqueda.
    Unload frmFind
End Sub

Private Sub cmdFind_Click()
    ' Asigna la cadena de texto a una variable pública.
    gFindString = txtFind.Text
    FindIt
End Sub

Private Sub Form_Load()
    ' Deshabilita el botón de buscar - todavía no hay texto que buscar.
    cmdFind.Enabled = False
    ' Lee la variable pública y establece el botón de opción.
    optDirection(gFindDirection).Value = 1
End Sub

Private Sub optDirection_Click(Index As Integer)
    ' Asigna un valor a la variable pública.
    gFindDirection = Index
End Sub

Private Sub txtFind_Change()
    ' Establece la variable pública.
    gFirstTime = True
    ' Si el cuadro de texto está vacío, deshabilita el botón de búsqueda.
    If txtFind.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

