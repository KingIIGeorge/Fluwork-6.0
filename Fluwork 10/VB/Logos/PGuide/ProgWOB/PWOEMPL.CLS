VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' >> Se ve mejor en la vista de Módulo completo. <<
'
' Almacenamiento para el número de ID de depuración.
Private mlngDebugID As Long
Implements IDebug

' Propiedades de la clase Employee.
Public Name As String
Public Salary As Double

' Datos privados de la propiedad ID de una sola escritura.
Private mstrID As String

Property Get ID() As String
    ID = mstrID
End Property
'
' La primera vez que se establece la propiedad ID, también se establece la
'   variable estática Boolean. La siguientes llamadas no hacen nada.
'   (Sería mejor desencadenar un error, en su lugar).
Property Let ID(strNew As String)
    Static blnAlreadySet As Boolean
    If Not blnAlreadySet Then
        blnAlreadySet = True
        mstrID = strNew
    End If
End Property

Private Sub Class_Initialize()
    mlngDebugID = DebugInit(Me)
End Sub

Private Sub Class_Terminate()
    DebugTerm Me
End Sub

' -------- Implementación de IDebug  --------
'
' IDebug.DebugID ofrece la manera de separar objetos.
'   Es requerido por los procedimientos de depuración
'   DebugInit, DebugTerm,
'   y DebugShow, declarados en
'   modFriend.
'
Private Property Get IDebug_DebugID() As Long
    IDebug_DebugID = mlngDebugID
End Property


