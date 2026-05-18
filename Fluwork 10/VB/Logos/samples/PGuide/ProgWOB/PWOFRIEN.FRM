VERSION 5.00
Begin VB.Form frmFriends 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amigos pasando tipos definidos por el usuario"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtString 
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Top             =   2640
      Width           =   3915
   End
   Begin VB.TextBox txtLong 
      Height          =   285
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   3
      Top             =   2160
      Width           =   1635
   End
   Begin VB.TextBox txtInteger 
      Height          =   285
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdMethod 
      Caption         =   "&Método Friend"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdProperty 
      Caption         =   "&Propiedad Friend"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo &String:"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo &Long:"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo &Integer:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   $"PWOFrien.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmFriends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Demuestra las propiedades Friend y los 
'   métodos pasando UDTs entre objetos.

Private mtc1 As TestClass
Private mtc2 As TestClass
    
' Utiliza propiedades para asignar y tener acceso a una UDT.
Private Sub cmdProperty_Click()
    ' El método SetDemoParts asigna
    '   el contenido de los cuadros de texto al 
    '   tipo definido por el usuario en el primer 
    '   objeto TestClass, así habrá algo 
    '   que pasar al segundo 
    '   objeto TestClass.
    Call mtc1.SetDemoParts(CInt("0" & txtInteger), _
        CLng("0" & txtLong), txtString)
    '
    ' Muestra los elementos UDT del primer 
    '   objeto TestClass antes de pasarlo.
    Call mtc1.ShowDemo("Pasando un UDT usando una propiedad", "Pasar desde:")
    '
    ' Asigna directamente la UDT del 
    '   primer objeto TestClass a la 
    '   UDT del segundo objeto TestClass,
    '   utilizando la propiedad Demo.
    mtc2.Demo = mtc1.Demo
    '
    ' Muestra los elementos UDT del segundo 
    '   objeto TestClass.
    Call mtc2.ShowDemo("Pasando un UDT usando una propiedad", "Pasado a:")
    '
    ' Cuando el procedimiento termina, tc1 y tc2
    '   pierden su alcance y los objetos TestClass
    '   terminan.
End Sub

' Utiliza métodos para asignar y tener acceso a una UDT.
Private Sub cmdMethod_Click()
    ' El método SetDemoParts asigna
    '   el contenido de los cuadros de texto al 
    '   tipo definido por el usuario en el primer 
    '   objeto TestClass, así habrá algo 
    '   que pasar al segundo 
    '   objeto TestClass.
    Call mtc1.SetDemoParts(CInt("0" & txtInteger), _
        CLng("0" & txtLong), txtString)
    '
    ' Muestra los elementos UDT del primer 
    '   objeto TestClass antes de pasarlo.
    Call mtc1.ShowDemo("Pasando un UDT usando un método", "Pasar desde:")
    '
    ' El método GetDemo del primer 
    '   objeto TestClass devuelve la UDT,
    '   que es pasada al método SetDemo
    '   del segundo objeto 
    '   TestClass.
    Call mtc2.SetDemo(mtc1.GetDemo)
    '
    ' Muestra los elementos UDT del segundo 
    '   objeto TestClass.
    Call mtc2.ShowDemo("Pasando un UDT usando un método", "Pasar a:")
    '
    ' Cuando el procedimiento termina, tc1 y tc2
    '   pierden su alcance y los objetos TestClass
    '   terminan.
End Sub

Private Sub Form_Load()
    ' Crea objetos TestClass.
    Set mtc1 = New TestClass
    Set mtc2 = New TestClass
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Libera los recursos del formulario.
    Set frmFriends = Nothing
End Sub

Private Sub txtInteger_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57     ' Permite los dígitos.
        Case 8      ' Permite el carácter de retroceso.
        Case Else   ' Suprime todos los demás.
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLong_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57     ' Permite los dígitos.
        Case 8      ' Permite el carácter de retroceso.
        Case Else   ' Suprime todos los demás.
            Beep
            KeyAscii = 0
    End Select
End Sub

