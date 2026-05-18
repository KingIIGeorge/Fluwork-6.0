Attribute VB_Name = "Module2"
' Obliga a tratar aquí todos los errores en tiempo de ejecución.
Sub DisplayErrorMessageBox()
    Dim Msg As String
    Select Case Err
        Case conMCIErrCannotLoadDriver
            Msg = "Error al cargar el controlador del dispositivo."
        Case conMCIErrDeviceOpen
            Msg = "El dispositivo no está abierto o no es conocido."
        Case conMCIErrInvalidDeviceID
            Msg = "Id. no válido de dispositivo"
        Case conMCIErrInvalidFile
            Msg = "Nombre de archivo no válido."
        Case conMCIErrUnsupportedFunction
            Msg = "Acción no disponible en este dispositivo."
        Case Else
            Msg = "Error desconocido (" + Str$(Err) + ")."
    End Select

    MsgBox Msg, 48, conMCIAppTIitle
End Sub

' Este procedimiento permite procesar cualquier evento de Windows.
' Esto puede ser necesario para resolver los problemas de 
' sincronización con los eventos de Windows.
' Este procedimiento puede emplearse asimismo para forzar un 
' retardo en el procesamiento.
Sub WaitForEventsToFinish(NbrTimes As Integer)
    Dim i As Integer

    For i = 1 To NbrTimes
        dummy% = DoEvents()
    Next i
End Sub


