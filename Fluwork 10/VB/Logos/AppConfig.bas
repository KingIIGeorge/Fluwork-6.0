Attribute VB_Name = "AppConfig"
Option Explicit

Public Function ReadDatabasePath() As String
    Dim fileNumber As Integer
    Dim value As String

    On Error GoTo ReadError

    fileNumber = FreeFile
    Open pconfig For Input As #fileNumber
    Line Input #fileNumber, value
    Close #fileNumber

    ReadDatabasePath = value
    Exit Function

ReadError:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    ReadDatabasePath = ""
End Function

Public Function ReadBaseValue(ByVal databasePath As String) As String
    Dim fileNumber As Integer
    Dim value As String

    On Error GoTo ReadError

    fileNumber = FreeFile
    Open databasePath + "\base.dat" For Input As #fileNumber
    Line Input #fileNumber, value
    Close #fileNumber

    ReadBaseValue = value
    Exit Function

ReadError:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    ReadBaseValue = ""
End Function

Public Function DataFilePath(ByVal fileName As String) As String
    DataFilePath = Trim$(dbpath) & "\" & fileName
End Function
