Attribute VB_Name = "Errors"
Option Explicit
Const mnErrDeviceUnavailable = 68
Const mnErrDiskNotReady = 71
Const mnErrDeviceIO = 57
Const mnErrDiskFull = 61
Const mnErrBadFileName = 64
Const mnErrBadFileNameOrNumber = 52
Const mnErrPathDoesNotExist = 76
Const mnErrBadFileMode = 54
Const mnErrFileAlreadyOpen = 55
Const mnErrInputPastEndOfFile = 62
Function FileErrors() As Integer
    Dim intMsgType As Integer
    Dim strMsg As String
    Dim intResponse As Integer
    ' Valor devuelto    Significado
    ' 0                 Continuar
    ' 1                 Continuar con la siguiente instrucción
    ' 2                 Error no recuperable
    ' 3                 Error no reconocido
    intMsgType = vbExclamation
    Select Case Err.Number
        Case mnErrDeviceUnavailable               ' Error 68
            strMsg = "El dispositivo no está disponible"
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrDiskNotReady                    ' Error 71
            strMsg = "Introduzca correctamente el disco en la unidad."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrDeviceIO                        ' Error 57
            strMsg = "Error interno de disco."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrDiskFull                        ' Error 61
            strMsg = "Disco lleno. ¿Desea continuar?"
            intMsgType = vbExclamation + vbAbortRetryIgnore
        Case mnErrBadFileName, mnErrBadFileNameOrNumber ' Error 64 y 52
            strMsg = "Nombre de archivo no válido."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrPathDoesNotExist                ' Error 76
            strMsg = "La ruta no existe."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrBadFileMode                     ' Error 54
            strMsg = "Imposible abrir el archivo para este tipo de acceso."
        Case mnErrFileAlreadyOpen                 ' Error 55
            strMsg = "El archivo ya está abierto."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrInputPastEndOfFile              ' Error 62
            strMsg = "Este archivo tiene un marcador no estándar de fin de archivo, "
            strMsg = strMsg & "o se ha intentado leer más allá "
            strMsg = strMsg & "del marcador de final de archivo."
            intMsgType = vbExclamation + vbAbortRetryIgnore
        Case Else
            FileErrors = 3
            Exit Function
    End Select
    intResponse = MsgBox(strMsg, intMsgType, "Error de disco")
    Select Case intResponse
        Case 1, 4       ' Botones Aceptar y Reintentar.
            FileErrors = 0
        Case 2, 5       ' Botones Cancelar e Ignorar.
            FileErrors = 1
        Case 3          ' Botón Abortar.
            FileErrors = 2
        Case Else
            FileErrors = 3
    End Select
End Function

