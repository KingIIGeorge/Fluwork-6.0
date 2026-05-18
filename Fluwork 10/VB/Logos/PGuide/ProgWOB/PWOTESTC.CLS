VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' >> Se ve mejor en la vista de Módulo completo. <<
'
' Almacenamiento para el número de ID de depuración.
Private mlngDebugID As Long
Implements IDebug

' Almacenamiento privado para el tipo definido
'   por el usuario (UDT) que se pasa entre los
'   objetos. El UDT tiene que ser declarado
'   Public en un módulo estándar (vea
'   modFriends). (En este ejemplo, mDemo
'   sirve como almacenamiento para la
'   propiedad Demo y para la pareja de función y
'   subrutina GetDemo/SetDemo.)
Private mDemo As udtDEMO

' Propiedad Demo:  Los Property Get y Let
'   de la propiedad Demo
'   muestran cómo los procedimientos de propiedad
'   declarados como Friend pueden aceptar y devolver
'   una variable UDT.
Friend Property Get Demo() As udtDEMO
    Demo = mDemo
End Property
' Observe que el argumento del Property
'   Let tiene que ser ByRef para una UDT.
Friend Property Let Demo(NewDemo As udtDEMO)
    mDemo = NewDemo
End Property

' GetDemo y SetDemo muestran cómo las funciones
'   Friend pueden
'   devolver UDTs, y cómo las UDTs pueden ser
'   argumentos de procedimientos Friend.
Friend Function GetDemo() As udtDEMO
    GetDemo = mDemo
End Function
' Observe que un argumento UDT en un procedimiento
'   tiene que ser declarado ByRef.
Friend Sub SetDemo(NewDemo As udtDEMO)
    mDemo = NewDemo
End Sub

' SetDemoParts es un método de ayuda que
'   llena los elementos
'   de la UDT, para que haya algo que pasar.
'   En un programa real, los elementos
'   de la UDT podrían llenarse desde el código
'   en el objeto TestClass.
Public Sub SetDemoParts(ByVal A As Integer, _
        ByVal B As Long, ByVal C As String)
    mDemo.intA = A
    mDemo.lngB = B
    mDemo.strC = C
End Sub

' ShowDemo presenta los elementos de la
'   UDT, de modo que puedan ser vistos
'   después de pasar la UDT al segundo
'   objeto.
Public Sub ShowDemo(ByVal Title As String, _
        ByVal Direction As String)
        
    MsgBox Direction & "  " _
        & GetDebugString(Me) & vbCrLf & vbCrLf _
        & mDemo.intA & vbCrLf _
        & mDemo.lngB & vbCrLf _
        & mDemo.strC, , Title
End Sub

Private Sub Class_Initialize()
    mlngDebugID = DebugInit(Me)
End Sub

Private Sub Class_Terminate()
    DebugTerm Me
End Sub

' -------- Implementación de IDebug --------
'
' IDebug.DebugID ofrece la manera de
'   distinguir los objetos. Es
'   requerido por los procedimientos de depuración
'   DebugInit, DebugTerm y DebugShow
'   declarados en modFriend.
'
Private Property Get IDebug_DebugID() As Long
    IDebug_DebugID = mlngDebugID
End Property

