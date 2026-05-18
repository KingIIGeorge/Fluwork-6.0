VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Módulo de clase de la aplicación de ejemplo VCR
' Actúa como módulo lógico para que VCR almacene
' datos y siga la actividad de grabación
'**********************************************

Option Explicit

Public Channel As Integer   'Canal de grabación
Public StartRecording As Variant    'Hora de inicio
Public StopRecording As Variant     'Hora de terminación
Public Enabled As Boolean   'Modo de grabación

Private Sub Class_Initialize()
    ' Inicializa las propiedades de la clase
    Enabled = True
    Channel = 3
    StartRecording = Empty
    StopRecording = Empty
End Sub
