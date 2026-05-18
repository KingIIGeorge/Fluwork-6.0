VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmallBusiness3"
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

' La colección privada Employees almacena los
'   objetos Employee.
Private emps As New Employees

' En lugar de declarar una variable Public
'   Employees As New Employees, es una mejor
'   práctica de programación hacer que la colección
'   sea privada, y crear una propiedad de sólo lectura
'   que devuelva la colección Employees.
'   De este modo, la colección
'   no puede ser liberada accidentalmente
'   estableciendo la variable pública a
'   Nothing.
Public Property Get Employees() As Employees
    Set Employees = emps
End Property

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


