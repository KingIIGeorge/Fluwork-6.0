Attribute VB_Name = "AppState"
Option Explicit

Public Sub ClearCurrentFicha()
    tmpficha = 0
End Sub

Public Sub SetFichaTouched(ByVal value As Boolean)
    touchedreally = value
End Sub

Public Sub SetCurrentFicha(ByVal ficha As Long)
    tmpficha = ficha
End Sub

Public Sub SelectFicha(ByVal ficha As Long)
    tmpficha = ficha
    touchedreally = True
End Sub

Public Sub SelectSingleFicha(ByVal ficha As Long)
    SelectFicha ficha
    onlyone = True
End Sub

Public Sub ClearSearchResults()
    tmpficha = 0
    showres = False
End Sub

Public Sub SetSearchResultsVisible(ByVal value As Boolean)
    showres = value
End Sub

Public Sub SetSingleResultMode(ByVal value As Boolean)
    onlyone = value
End Sub

Public Sub SetConfirmationFilter(ByVal value As String)
    conpre = value
End Sub
