VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employees"
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

' Colección privada que realiza todo el
'   trabajo.
Private colEmployees As New Collection

' El método NewEnum permite el uso de For
'    Each con la clase de colección
'    Employees. Para obtener su explicación,
'    vea "Creación de su propia clase de
'    colección: la casa de ladrillos"
'    dentro de los Libros en pantalla.
'
' Para que NewEnum funcione correctamente, su
'   Id de procedimiento tiene que ser -4.  Para ver
'   esto, seleccione Atributos del procedimiento
'   en el menú Herramientas. En el cuadro Nombre,
'   seleccione Item, después haga clic en el botón
'   Avanzados.  Puede ver que el ID de procedimiento
'   ha sido establecido a -4.  (También está activado
'   "Ocultar este miembro", de forma que el método
'   NewEnum está oculto en el
'   Examinador de objetos).
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' Delega en el método _NewEnum
    '   del objeto de la colección privada.
    Set NewEnum = colEmployees.[_NewEnum]
End Function

' El método Add crea un nuevo Employee en
'   la colección.
'
Public Function Add(ByVal Name As String, _
                ByVal Salary As Double) As Employee
    Dim empNew As New Employee
    Static intEmpNum As Integer
    ' El uso de With acorta las referencias a las propiedades (.ID en lugar de empNew.ID)
    ' y acelera la ejecución.
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
    ' Es una buena costumbre devolver la referencia del nuevo miembro.
    Set Add = empNew
End Function

' Para ver cómo el método Item se convierte en el
'   método predeterminado de la colección Employees,
'   seleccione Atributos del procedimiento
'   en el menú Herramientas.  En el cuadro Nombre,
'   seleccione Item, después haga clic en el botón
'   Avanzados. Puede ver que (Predeterminado)
'   ha sido seleccionado en el cuadro del ID
'   de procedimiento.
Public Function Item(ByVal Index As Variant) As Employee
    Set Item = colEmployees.Item(Index)
End Function

Public Function Count() As Long
    Count = colEmployees.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    colEmployees.Remove Index
End Sub

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


