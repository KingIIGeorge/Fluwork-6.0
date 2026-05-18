VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotifyMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' NotifyMe implementa ICoffeeNotify para que
'   su método de llamada de retorno (CoffeeReady)
'   pueda llamarse con vinculación temprana.
Implements ICoffeeNotify

' Almacenamiento para la implementación de NotifyID.
Private mlngNotifyID As Long

' ICoffeeNotify.CoffeeReady es llamado por el objeto
' ============= -----------     CoffeeMonitor2
'   de Coffee2. Esta implementación agrega la hora de
'   la llamada a un cuadro de lista, eliminando la
'   entrada más antigua si hay más de diez.
'
Private Sub ICoffeeNotify_CoffeeReady()
    With Form1.lstCallBacks
        .AddItem Format$(Now, "ddd hh:mm:ss"), 0
        If .ListCount > 10 Then .RemoveItem 10
    End With
End Sub

' ICoffeeNotify.NotifyID. Esta propiedad contiene la 
' ============= --------       clave del objeto a
'   notificar en la matriz de objetos cliente de 
'   CoffeeMonitor2. La implementación cliente sólo
'   almacena y devuelve el valor.
'
Private Property Let ICoffeeNotify_NotifyID(ByVal RHS As Long)
    mlngNotifyID = RHS
End Property
'
Private Property Get ICoffeeNotify_NotifyID() As Long
    ICoffeeNotify_NotifyID = mlngNotifyID
End Property
