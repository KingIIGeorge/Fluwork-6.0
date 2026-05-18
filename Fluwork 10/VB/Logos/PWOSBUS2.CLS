VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmallBusiness2"
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

' El objeto Collection privado es algo más
'   robusto que el objeto Collection público,
'   pero también tiene problemas.
Private colEmployees As New Collection

' Problema 1: La colección es privada,
'   de modo que el método Add tiene que formar parte de
'   la clase SmallBusiness2. Puesto que
'   puede haber muchas clases de objetos en una
'   empresa, puede haber muchos métodos de dicho tipo,
'   como AddProduct.
'
Public Function AddEmployee(ByVal Name As String, ByVal Salary As Double) As Employee
    Dim empNew As New Employee
    Static intEmpNum As Integer
    ' El uso de With acorta las referencias de las propiedades
    ' (.ID en lugar de empNew.ID) y acelera la ejecución.
    With empNew
        ' Genera un ID único para el nuevo empleado.
        intEmpNum = intEmpNum + 1
        .ID = "E" & Format$(intEmpNum, "00000")
        .Name = Name
        .Salary = Salary
        ' Agrega la referencia del objeto Employee a la colección, utilizando su
        ' propiedad ID como clave.
        colEmployees.Add empNew, .ID
    End With
    ' El método Add debería devolver una referencia al
    '   nuevo objeto.
    Set AddEmployee = empNew
End Function

' Problema 1, continuación:  También habrá
'   funciones Count para todas las clases de
'   objetos y también
'   métodos Delete, etc..  Incluso solamente
'   con los objetos Employee, la clase
'   SmallBusiness2 queda muy
'   complicada.
'
Public Function EmployeeCount() As Long
    EmployeeCount = colEmployees.Count
End Function

Public Sub DeleteEmployee(ByVal Index As Variant)
    colEmployees.Remove Index
End Sub

' Problema 2: El método de Employees no puede ser
'   utilizado con For Each. Podría hacer que funcione
'   con For Each devolviendo una
'   referencia a la colección privada,
'   pero entonces ya no sería una colección privada,
'   y volveríamos a la
'   Casa de paja.
'
Public Function Employees(ByVal Index As Variant) As Employee
    Set Employees = colEmployees.Item(Index)
End Function

' Problema 3: Aún es posible
'   codificar errores, esta vez dentro del
'   mucho más complicado objeto SmallBusiness,
'   agregando objetos no válidos a la
'   colección.
Public Sub Trouble()
    Dim X As New Employee
    ' Puesto que el objeto Collection acepta
    '   una referencia a cualquier objeto, aún es
    '   posible agregar un objeto Employee no
    '   inicializado para tener un
    '   error de código.
    colEmployees.Add X
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


