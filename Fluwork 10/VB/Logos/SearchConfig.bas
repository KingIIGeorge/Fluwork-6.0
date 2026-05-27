Attribute VB_Name = "SearchConfig"
Option Explicit

Public Function SearchApiBaseUrl() As String
    Dim configuredUrl As String

    configuredUrl = ReadSearchApiBaseUrlFromFile()
    If Trim$(configuredUrl) <> "" Then
        SearchApiBaseUrl = configuredUrl
    Else
        SearchApiBaseUrl = "http://127.0.0.1:5050"
    End If
End Function

Private Function ReadSearchApiBaseUrlFromFile() As String
    On Error GoTo ErrHandler

    Dim fileNumber As Integer
    Dim configPath As String
    Dim value As String

    configPath = App.Path & "\fluwork_search_url.txt"
    If Dir$(configPath) = "" Then Exit Function

    fileNumber = FreeFile
    Open configPath For Input As #fileNumber
    Line Input #fileNumber, value
    Close #fileNumber

    ReadSearchApiBaseUrlFromFile = Trim$(value)
    Exit Function

ErrHandler:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    ReadSearchApiBaseUrlFromFile = ""
End Function
