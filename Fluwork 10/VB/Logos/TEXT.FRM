VERSION 5.00
Begin VB.Form frmText 
   Caption         =   "Propiedades del cuadro de texto"
   ClientHeight    =   4455
   ClientLeft      =   930
   ClientTop       =   1290
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6915
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Restablecer"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fraInsert 
      Caption         =   "Establecer el punto de inserción"
      Height          =   2295
      Left            =   180
      TabIndex        =   5
      Top             =   1200
      Width           =   3675
      Begin VB.OptionButton optText 
         Caption         =   "Insertar te&xto"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "&Seleccionar todo el texto"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   2475
      End
      Begin VB.OptionButton optInsert 
         Caption         =   "&Punto de inserción después del 5º carácter"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3435
      End
      Begin VB.OptionButton optEnd 
         Caption         =   "P&unto de inserción al final"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2475
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "&Valores predeterminados"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.TextBox txtDisplay 
      Height          =   375
      HideSelection   =   0   'False
      Left            =   180
      TabIndex        =   0
      Text            =   "La propiedad MultiLine es False en este ejemplo"
      Top             =   480
      Width           =   3675
   End
   Begin VB.TextBox txtMulti 
      Height          =   615
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "text.frx":0000
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblHelp 
      Caption         =   "Seleccione una opción para ver su efecto en el primer cuadro de texto."
      Height          =   735
      Left            =   4200
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me   ' Descarga este formulario.
End Sub

Private Sub cmdReset_Click()
    ' restablece el texto original
    txtDisplay.Text = "La propiedad MultiLine es False en este ejemplo"
    ' restablece los botones de opción a su estado predeterminado
    optDefault.Value = True
End Sub

Private Sub optDefault_Click()
    ' coloca el punto de inserción al principio
    txtDisplay.SelStart = 0
    
    ' establece el enfoque sobre el cuadro de texto para que podamos ver
    ' el resultado
    txtDisplay.SetFocus
End Sub

Private Sub optEnd_Click()
    ' obtiene la longitud de la cadena y coloca el
    ' punto de inserción al final
    txtDisplay.SelStart = Len(txtDisplay.Text)
    
    ' establece el enfoque sobre el cuadro de texto para que podamos ver
    ' el resultado
    txtDisplay.SetFocus
End Sub

Private Sub optInsert_Click()
    ' coloca el punto de inserción después del quinto carácter
    txtDisplay.SelStart = 5
        
    ' establece el enfoque sobre el cuadro de texto para que podamos ver
    ' el resultado
    txtDisplay.SetFocus
End Sub

Private Sub optSelect_Click()
    ' coloca el punto de inserción al principio
    txtDisplay.SelStart = 0
    ' obtiene la longitud de la cadena y
    ' selecciona dicho número de caracteres
    txtDisplay.SelLength = Len(txtDisplay.Text)
    
    ' establece el enfoque sobre el cuadro de texto para que podamos ver
    ' el resultado
    txtDisplay.SetFocus
End Sub

Private Sub optText_Click()
    ' inserta "NEW STRING" en el punto de inserción
    ' si hubiera texto seleccionado, será reemplazado
    txtDisplay.SelText = "TEXTO NUEVO"
    
    ' establece el enfoque sobre el cuadro de texto para que podamos ver
    ' el resultado
    txtDisplay.SetFocus
End Sub

