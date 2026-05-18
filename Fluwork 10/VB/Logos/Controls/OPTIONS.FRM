VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Botones de opción"
   ClientHeight    =   3825
   ClientLeft      =   3255
   ClientTop       =   2400
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5040
   Begin VB.OptionButton opt686 
      Caption         =   "P&entium Pro"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton opt586 
      Caption         =   "&Pentium"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton opt486 
      Caption         =   "&486"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Frame fraSystem 
      Caption         =   "Sistema operativo"
      Height          =   1455
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton optWin95 
         Caption         =   "&Windows 95"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optWinNT 
         Caption         =   "Windows &NT"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label lblHelp 
      Caption         =   "Seleccione un procesador y un sistema operativo"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declara dos variables de cadena para almacenar los títulos
Dim strComputer As String
Dim strSystem As String

Sub DisplayCaption()
    ' concatena el título con las dos
    ' variables de cadena.
    lblDisplay.Caption = "Ha seleccionado un " & _
     strComputer & " ejecutando " & strSystem
End Sub

Private Sub cmdClose_Click()
    Unload Me   'descarga el formulario
End Sub

Private Sub Form_Load()
    ' invoca un evento Click en las opciones predeterminadas
    ' para actualizar el título de la etiqueta
    opt486_Click
    optWin95_Click
End Sub

Private Sub opt486_Click()
    ' asigna un valor a la primera variable de cadena
    strComputer = "486"
    ' llama a la subrutina
    Call DisplayCaption
End Sub


Private Sub opt586_Click()
    ' asigna un valor a la primera variable de cadena
    strComputer = "Pentium"
    ' llama a la subrutina
    Call DisplayCaption
End Sub


Private Sub opt686_Click()
    ' asigna un valor a la primera variable de cadena
    strComputer = "Pentium Pro"
    ' llama a la subrutina
    Call DisplayCaption
End Sub


Private Sub optWin95_Click()
    ' asigna un valor a la segunda variable de cadena
    strSystem = "Windows 95"
    ' llama a la subrutina
    Call DisplayCaption
End Sub


Private Sub optWinNT_Click()
    ' asigna un valor a la segunda variable de cadena
    strSystem = "Windows NT"
    ' llama a la subrutina
    Call DisplayCaption
End Sub


