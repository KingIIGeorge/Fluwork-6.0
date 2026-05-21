Attribute VB_Name = "GlobalTypes"
Option Explicit

Public Type tpersona
    persona1 As String * 30
    persona2 As String * 30
    persona3 As String * 30
    persona4 As String * 30
    persona5 As String * 30
    persona6 As String * 30
    persona7 As String * 30
    persona8 As String * 30
    persona9 As String * 30
    persona10 As String * 30
    persona11 As String * 30
    persona12 As String * 30
    persona13 As String * 30

    persona14 As String * 200
    persona15 As String * 200
    persona16 As String * 200
    persona17 As String * 200
    persona18 As String * 200
    persona19 As String * 200
    persona20 As String * 200
    persona21 As String * 200
    persona22 As String * 200
    persona23 As String * 200
    persona24 As String * 200
    persona25 As String * 200
    persona26 As String * 200
End Type

Public Type Tregistro
    ficha As String * 10
    fechaingreso As String * 10
    fechaegreso As String * 10
    estado As String * 10
    fullname As String * 50
    telefono As String * 15
    adjuntos As String * 1024
    problema As String * 1024
    solucion As String * 1024
    presupuesto As String * 10
    precio As String * 10
    atendidopor As String * 50
    tecnico As String * 50
    modelo As String * 50
    nserie As String * 50
    direccion As String * 200
    email As String * 75
    llamareldia As String * 30
    controladopor As String * 50
    avisadoeldia As String * 30
    avisadopor As String * 30
    confirmacion As String * 30
End Type

Public Type Tindexregistro
    ficha As String * 10
    fullname As String * 50
    telefono As String * 15
    modelo As String * 50
    fecha As String * 10
    estado As String * 10
    tecnico As String * 50
    confirmacion As String * 30
End Type

Public Type Testados
    txt As String
End Type
