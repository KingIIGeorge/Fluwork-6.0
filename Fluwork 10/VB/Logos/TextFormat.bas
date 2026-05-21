Attribute VB_Name = "TextFormat"
Option Explicit

Public Function EnterToBr(ByVal value As String) As String
Dim resultado As String
Dim i As Long
Dim pos As Long
resultado = ""
pos = 0

For i = 1 To Len(value)
If Mid(value, i, 1) = vbCr Or i = Len(value) Then
If pos = 0 Then
resultado = Mid(value, 1, i - 1) & "<br>"
ElseIf i = Len(value) Then
resultado = resultado & Mid(value, pos + 1, i - pos) & "<br>"
Else
resultado = resultado & Mid(value, pos + 1, i - pos - 1) & "<br>"
End If
pos = i + 1
End If
Next i
EnterToBr = resultado
End Function

Public Function PipeToBr(ByVal value As String) As String
Dim resultado As String
Dim i As Long
Dim pos As Long
Dim npipe As Byte

npipe = 0
resultado = ""
pos = 0

For i = 1 To Len(value)
If Mid(value, i, 1) = "|" Or i = Len(value) Then
npipe = npipe + 1
    If pos = 0 Then
    resultado = Mid(value, 1, i - 1) & "   "
    ElseIf i = Len(value) Then
    resultado = resultado & Mid(value, pos, i - pos + 1) & "<BR>"
    Else
    If npipe Mod 2 = 0 Then
    resultado = resultado & Mid(value, pos, i - pos) & "<BR>"
    Else
    resultado = resultado & Mid(value, pos, i - pos) & "   "
    End If
    End If
    pos = i + 1
End If
Next i
PipeToBr = resultado
End Function
