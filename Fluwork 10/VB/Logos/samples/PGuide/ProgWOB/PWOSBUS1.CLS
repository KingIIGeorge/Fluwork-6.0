VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmallBusiness1"
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

' El objeto Collection público es fácil de
'   declarar, pero trae complicaciones
'   como se ve en frmStraw.
Public Employees As New Collection

Property Get NewEmployeeID() As String
    Static intEmpNum As Integer
    intEmpNum = intEmpNum + 1
    NewEmployeeID = "E" & Format$(intEmpNum, "00000")
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


