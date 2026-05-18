Attribute VB_Name = "modATM"
Option Explicit

' Separa la variable
Public I As Long

' Tabla de conversión de moneda para dólares americanos
Public ConversionTable As Variant

' API de soporte de sonido de alto nivel
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = &H1     ' Reproduce de forma asíncrona
Global Const SND_NODEFAULT = &H2 ' No usar el sonido predeterminado
Global Const SND_MEMORY = &H4    ' lpszSoundName apunta al archivo de memoria

Global SoundBuffer() As Byte

' Puntero del ratón sobre el botón de comando.
Dim curSelect As StdPicture

Sub ConversionTable_Initialize()
    ' Actualizado el 14/08/1996.
    ConversionTable = Array(1@, 5.0776@, 1.4861@, 1519.28@, 126.1@, 107.95@)
End Sub

Sub BeginPlaySound(ByVal ResourceId As Integer)
    SoundBuffer = LoadResData(ResourceId, "ATM_SOUND")
    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub

Sub EndPlaySound()
    sndPlaySound ByVal vbNullString, 0&
End Sub

Sub Cursor_Initialize()
    Set curSelect = LoadResPicture(1, vbResCursor)
End Sub

Sub SetCursor(Button As CommandButton)
    With Button
        .MousePointer = 99
        .MouseIcon = curSelect
    End With
End Sub
