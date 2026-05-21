Attribute VB_Name = "PathUtils"
Option Explicit

Public Function EnsureTrailingBackslash(ByVal value As String) As String
    If (Right(value, 1) = "\") Then
        EnsureTrailingBackslash = value
    Else
        EnsureTrailingBackslash = value & "\"
End If
End Function
