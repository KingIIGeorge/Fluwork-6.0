VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   1395
   ClientLeft      =   2655
   ClientTop       =   3585
   ClientWidth     =   5640
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
   ScaleWidth      =   5640
   Begin VB.Frame Frame1 
      Caption         =   "Dirección"
      Height          =   612
      Left            =   2340
      TabIndex        =   3
      Top             =   660
      Width           =   1995
      Begin VB.OptionButton optDirection 
         Caption         =   "A&bajo"
         Height          =   252
         Index           =   1
         Left            =   1020
         TabIndex        =   5
         ToolTipText     =   "Buscar hasta el final del documento"
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "&Arriba"
         Height          =   252
         Index           =   0
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "Buscar hasta el inicio del documento"
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "&Coincidencia exacta"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Coincidir MAY/min"
      Top             =   780
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   500
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Texto para buscar"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   372
      Left            =   4500
      TabIndex        =   7
      ToolTipText     =   "Volver al Bloc de notas"
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "B&uscar"
      Default         =   -1  'True
      Height          =   372
      Left            =   4500
      TabIndex        =   6
      ToolTipText     =   "Iniciar búsqueda"
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Bu&scar:"
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
'*** Cuadro de diálogo Buscar para búsquedas de texto.                  ***
'*** Creado para la aplicación de ejemplo Bloc de notas MDI             ***
'*** Utiliza las variables públicas gFindCase (alterna la distinción    ***
'*** entre mayúsculas y minúsculas); gFindString (texto que se          ***
'*** busca); gFindDirection (alterna la dirección de la                 ***
'*** búsqueda); gFirstTime (alterna el inicio desde el principio        ***
'*** del texto)                                                         ***
'**************************************************************************

Option Explicit

Private Sub chkCase_Click()
    ' Asigna un valor a la variable pública.
    gFindCase = chkCase.Value
End Sub

Private Sub cmdCancel_Click()
    ' Guarda los valores en las variables públicas.
    gFindString = Text1.Text
    gFindCase = chkCase.Value
    ' Descarga el cuadro de diálogo de búsqueda.
    Unload frmFind
End Sub

Private Sub cmdFind_Click()
    ' Asigna la cadena de texto a una variable pública.
    gFindString = Text1.Text
    FindIt
End Sub

Private Sub Form_Load()
    ' Deshabilita el botón de búsqueda - todavía no hay texto que buscar.
    cmdFind.Enabled = False
    ' Lee la variable pública y establece el botón de opción.
    optDirection(gFindDirection).Value = 1
End Sub

Private Sub optDirection_Click(index As Integer)
    ' Asigna un valor a la variable pública.
    gFindDirection = index
End Sub

Private Sub Text1_Change()
    ' Establece la variable pública.
    gFirstTime = True
    ' Si el cuadro de texto está vacío, deshabilita el botón de búsqueda.
    If Text1.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

