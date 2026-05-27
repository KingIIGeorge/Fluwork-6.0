Attribute VB_Name = "RecordManager"
Option Explicit

Public Function LastFichaNumber() As Long
    Dim indexRecord As Tindexregistro
    Dim fileNumber As Integer
    Dim fileSize As Long

    fileNumber = FreeFile
    Open AppConfig.DataFilePath("index.dat") For Random Access Read As #fileNumber Len = Len(indexRecord)
    fileSize = LOF(fileNumber)

    If fileSize > 0 Then
        Get #fileNumber, fileSize / Len(indexRecord), indexRecord
        LastFichaNumber = Val(Trim$(indexRecord.ficha))
    Else
        LastFichaNumber = BASE
    End If

    Close #fileNumber
End Function

Public Function ReadIndexRecord(ByVal recordNumber As Long, ByRef outIndex As Tindexregistro) As Boolean
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open AppConfig.DataFilePath("index.dat") For Random Access Read As #fileNumber Len = Len(outIndex)
    Get #fileNumber, recordNumber, outIndex
    Close #fileNumber

    ReadIndexRecord = True
End Function

Public Function ReadFicha(ByVal ficha As Long, ByRef outRegistro As Tregistro) As Boolean
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open AppConfig.DataFilePath(Trim$(str$(ficha))) For Binary Access Read As #fileNumber Len = Len(outRegistro)
    Get #fileNumber, , outRegistro
    Close #fileNumber

    ReadFicha = True
End Function

Public Sub BuildIndexFromRegistro(ByRef source As Tregistro, ByRef target As Tindexregistro)
    target.ficha = source.ficha
    target.fullname = StrConv(source.fullname, 1)
    target.telefono = source.telefono
    target.modelo = source.modelo
    target.fecha = source.fechaingreso
    target.estado = source.estado
    target.tecnico = source.tecnico
    target.confirmacion = source.confirmacion
End Sub

Public Function WriteFichaAndIndex(ByRef item As Tregistro) As Boolean
    Dim fileNumber As Integer
    Dim itemIndex As Tindexregistro

    fileNumber = FreeFile
    Open AppConfig.DataFilePath(Trim$(item.ficha)) For Random As #fileNumber Len = Len(item)
    Put #fileNumber, , item
    Close #fileNumber

    BuildIndexFromRegistro item, itemIndex

    fileNumber = FreeFile
    Open AppConfig.DataFilePath("index.dat") For Random As #fileNumber Len = Len(itemIndex)
    Put #fileNumber, Val(item.ficha) - BASE, itemIndex
    Close #fileNumber

    WriteFichaAndIndex = True

    On Error Resume Next
    Call SearchHttpClient.NotificarFichaActualizada(Val(item.ficha))
    On Error GoTo 0
End Function
