Attribute VB_Name = "StatusCatalog"
Option Explicit

Public Const STATUS_COUNT As Integer = 15

Public Sub LoadDefaultStatuses()
    estados(1).txt = "POR VER"
    estados(2).txt = "REPARANDO"
    estados(3).txt = "LISTA"
    estados(4).txt = "STD/BY"
    estados(5).txt = "ENTREGADA"
    estados(6).txt = "CHEQUEO"
    estados(7).txt = "REP.EXT."
    estados(8).txt = "PV EXT."
    estados(9).txt = "LISTA NR"
    estados(10).txt = "LISTA BRGS"
    estados(11).txt = "ENTREGAR"
    estados(12).txt = "PRESUP"
    estados(13).txt = "ANULADA"
    estados(14).txt = "DEPOSITO"
    estados(15).txt = "DIAGNOSTIC"
End Sub

Public Sub ApplyStatusSelection(ByVal statusIndex As Integer)
    If statusIndex < 1 Or statusIndex > STATUS_COUNT Then Exit Sub

    registro.estado = estados(statusIndex).txt

    Select Case statusIndex
        Case 1, 2, 4, 6, 7, 8, 12
            registro.fechaegreso = " "
            registro.precio = " "
        Case 3, 9, 10
            registro.fechaegreso = Trim(str(Date))
            registro.precio = " "
        Case 5
            registro.precio = Trim(str(Date))
        Case 11
            registro.precio = " "
        Case 13, 15
            registro.fechaegreso = ""
            registro.precio = ""
    End Select
End Sub

Public Function ColorForStatus(ByVal statusText As String) As Long
    Select Case Trim$(statusText)
        Case "POR VER"
            ColorForStatus = QBColor(11)
        Case "REPARANDO"
            ColorForStatus = QBColor(12)
        Case "LISTA"
            ColorForStatus = QBColor(10)
        Case "ENTREGADA"
            ColorForStatus = QBColor(8)
        Case "STD/BY"
            ColorForStatus = QBColor(13)
        Case "CHEQUEO"
            ColorForStatus = QBColor(14)
        Case "REP.EXT."
            ColorForStatus = QBColor(12)
        Case "PV EXT."
            ColorForStatus = QBColor(11)
        Case "LISTA NR"
            ColorForStatus = QBColor(10)
        Case "LISTA BRGS"
            ColorForStatus = QBColor(10)
        Case "PRESUP"
            ColorForStatus = QBColor(15)
        Case "ENTREGAR"
            ColorForStatus = QBColor(9)
        Case "ANULADA"
            ColorForStatus = QBColor(2)
        Case "DEPOSITO"
            ColorForStatus = QBColor(8)
        Case "DIAGNOSTIC"
            ColorForStatus = QBColor(14)
        Case Else
            ColorForStatus = -1
    End Select
End Function
