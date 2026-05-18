VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' El módulo de clase IShape es una "clase
'   abstracta" que define la interfaz IShape.
'   (Hablando con propiedad, no es una
'   clase abstracta:  Una clase verdaderamente
'   abstracta no se puede utilizar para crear objetos.
'   Esta clase, como todas las clase definidas en
'   Visual Basic, puede ser utilizada para crear
'   objetos, pero dichos objetos no serían
'   particularmente útiles porque, como
'   una clase realmente abstracta, hay miembros sin
'   implementación).

' El método TimeTest no tiene argumentos, y
'   su implementación siempre
'   está vacía esto es, no ejecuta
'   nada de código. Su único propósito es el de
'   ilustrar la asociación de tipos previa y tardía.
Public Sub TimeTest()
End Sub

Public Sub DrawToPictureBox(ByVal pb As PictureBox)
End Sub

