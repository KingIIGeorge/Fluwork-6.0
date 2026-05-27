Attribute VB_Name = "SearchHttpClient"
Option Explicit

Public Function BuscarClienteTexto(ByVal cliente As String, ByVal limit As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?cliente=" & UrlEncode(cliente) & "&limit=" & CStr(limit)
    BuscarClienteTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarClientePaginadoTexto(ByVal cliente As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?cliente=" & UrlEncode(cliente) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarClientePaginadoTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarEstadoTexto(ByVal estado As String, ByVal limit As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?estado=" & UrlEncode(estado) & "&limit=" & CStr(limit)
    BuscarEstadoTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarEstadoPaginadoTexto(ByVal estado As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?estado=" & UrlEncode(estado) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarEstadoPaginadoTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarTecnicoTexto(ByVal tecnico As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?tecnico=" & UrlEncode(tecnico) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarTecnicoTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarTelefonoTexto(ByVal telefono As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?telefono=" & UrlEncode(telefono) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarTelefonoTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarModeloTexto(ByVal modelo As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?modelo=" & UrlEncode(modelo) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarModeloTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarTecnicoEstadoTexto(ByVal tecnico As String, ByVal estado As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?tecnico=" & UrlEncode(tecnico) & "&estado=" & UrlEncode(estado) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarTecnicoEstadoTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarTecnicoConfirmacionTexto(ByVal tecnico As String, ByVal confirmacion As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?tecnico=" & UrlEncode(tecnico) & "&confirmacion=" & UrlEncode(confirmacion) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarTecnicoConfirmacionTexto = HttpGetText(url, responseText)
End Function

Public Function BuscarConfirmacionTexto(ByVal confirmacion As String, ByVal limit As Long, ByVal offset As Long, ByRef responseText As String) As Boolean
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/buscar_vb6?confirmacion=" & UrlEncode(confirmacion) & "&limit=" & CStr(limit) & "&offset=" & CStr(offset)
    BuscarConfirmacionTexto = HttpGetText(url, responseText)
End Function

Public Function NotificarFichaActualizada(ByVal ficha As Long) As Boolean
    Dim responseText As String
    Dim url As String

    url = SearchConfig.SearchApiBaseUrl() & "/api/reindex_ficha_vb6?ficha=" & CStr(ficha)
    NotificarFichaActualizada = HttpGetText(url, responseText)
End Function

Private Function HttpGetText(ByVal url As String, ByRef responseText As String) As Boolean
    On Error GoTo ErrHandler

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.SetTimeouts 1500, 1500, 5000, 5000
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then Exit Function

    responseText = http.ResponseText
    HttpGetText = True
    Exit Function

ErrHandler:
    responseText = ""
    HttpGetText = False
End Function

Private Function UrlEncode(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122
                result = result & ch
            Case 32
                result = result & "+"
            Case 0 To 127
                AppendEncodedByte result, code
            Case 128 To 2047
                AppendEncodedByte result, &HC0 Or (code \ 64)
                AppendEncodedByte result, &H80 Or (code And &H3F)
            Case Else
                AppendEncodedByte result, &HE0 Or (code \ 4096)
                AppendEncodedByte result, &H80 Or ((code \ 64) And &H3F)
                AppendEncodedByte result, &H80 Or (code And &H3F)
        End Select
    Next i

    UrlEncode = result
End Function

Private Sub AppendEncodedByte(ByRef result As String, ByVal value As Long)
    result = result & "%" & Right$("0" & Hex$(value And &HFF), 2)
End Sub
