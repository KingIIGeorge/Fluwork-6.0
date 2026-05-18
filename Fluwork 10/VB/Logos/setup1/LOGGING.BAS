Attribute VB_Name = "basLogging"
Option Explicit
Option Compare Text

'
' Módulo basLogging (sólo para 32 bits)
'
'Las rutinas de este módulo se usan para iniciar acciones,
'advertencias, notas y errores en el archivo de registro para
'desinstalar aplicaciones. Este archivo de registro lo usará
'la aplicación de desinstalación (ST5UNST.EXE) cuando el usuario
'decida desinstalar la aplicación (con el icono del Administrador
'de programas en Windows NT o con la función para Agregar o quitar
'programas en el panel de control de Windows 95).
'
'Las funciones están basadas en transacciones como "acciones".
'Cuando el programa de instalación comienza a procesar una nueva acción
'(una acción es algo que la utilidad para desinstalar aplicaciones
'debe deshacer), es necesario llamar a la función NewAction() con
'los parámetros apropiados para la acción
'(busque NewAction en este proyecto para ver cómo se forman
'los parámetros correctos para varias acciones).
'Cuando la acción se ha completado, se llama
'a la función CommitAction(), o si la acción no se
'ha completado correctamente, se llama a AbortAction()
'. Si se llama a CommitAction(), entonces la
'acción se registra en este momento, y la utilidad para
'desinstalar aplicaciones podrá deshacer la acción (por ejemplo,
'eliminar un archivo que ha copiado el instalador).
'
'La acciones pueden estar anidadas (por ejemplo, una acción
'para copiar un archivo puede tener acciones para crear directorios).
'Los errores, advertencias o notas registrarán en el
'archivo de registro la acción pendiente (si la hay). Incluso si
'se registra un error, es necesario que la acción siguiente
'se realice o se cancele. Vea los comentarios de cada función
'para obtener más detalles.
'

'La desinstalación de aplicaciones sólo es posible en aplicaciones de 32 bits

'Establezca esta constante a FALSE si no desea que aparezcan advertencias
'en el archivo de registro
Global Const fLOG_WARNINGS = True

'Constantes de clave para acciones globales
Global Const gstrKEY_PRIVATEFILE = "PrivateFile"
Global Const gstrKEY_TEMPFILE = "TempFile"
Global Const gstrKEY_SHAREDFILE = "SharedFile"
Global Const gstrKEY_SYSTEMFILE = "SystemFile"
Global Const gstrKEY_CREATEDIR = "CreateDir"
Global Const gstrKEY_PROGMANGROUP = "ProgManGroup"
Global Const gstrKEY_PROGMANITEM = "ProgManItem"
Global Const gstrKEY_SHELLFOLDER = "ShellFolder"
Global Const gstrKEY_SHELLLINK = "ShellLink"
Global Const gstrKEY_DLLSELFREGISTER = "DllSelfRegister"
Global Const gstrKEY_EXESELFREGISTER = "ExeSelfRegister"
Public Const gstrKEY_TLBREGISTER = "TLBRegister"
Global Const gstrKEY_REMOTEREGISTER = "RemoteRegister"
Global Const gstrKEY_REGKEY = "RegKey"
Global Const gstrKEY_REGVALUE = "RegValue"

'Errores de registro de VB5STKIT.DLL
Private Const LOGERR_SUCCESS = 0
Private Const LOGERR_INVALIDARGS = 1
Private Const LOGERR_OUTOFMEMORY = 2
Private Const LOGERR_EXCEEDEDCAPACITY = 3
Private Const LOGERR_WRITEERROR = 4
Private Const LOGERR_NOCURRENTACTION = 5
Private Const LOGERR_UNEXPECTED = 6
Private Const LOGERR_FILENOTFOUND = 7

'Severidad de errores de registro
Private Const LogErrOK = 1 ' Continuamos a pesar del error
Private Const LogErrFatal = 2 ' Es necesario terminar

'interfaces de SKIT432.DLL
Private Declare Function DllAbortAction Lib "VB5STKIT.DLL" Alias "AbortAction" () As Long
Private Declare Function DllAddActionNote Lib "VB5STKIT.DLL" Alias "AddActionNote" (ByVal lpszNote As String) As Long
Private Declare Function DllChangeActionKey Lib "VB5STKIT.DLL" Alias "ChangeActionKey" (ByVal lpszNewKey As String) As Long
Private Declare Function DllCommitAction Lib "VB5STKIT.DLL" Alias "CommitAction" () As Long
Private Declare Function fDllWithinAction Lib "VB5STKIT.DLL" Alias "fWithinAction" () As Long
Private Declare Function DllLogError Lib "VB5STKIT.DLL" Alias "LogError" (ByVal lpszERROR As String, ByVal lpszDURINGACTION As String, ByVal lpszErrMsg As String) As Long
Private Declare Function DllLogNote Lib "VB5STKIT.DLL" Alias "LogNote" (ByVal lpszNote As String) As Long
Private Declare Function DllLogWarning Lib "VB5STKIT.DLL" Alias "LogWarning" (ByVal lpszWARNING As String, ByVal lpszDURINGACTION As String, ByVal lpszWarningMsg As String) As Long
Private Declare Function DllNewAction Lib "VB5STKIT.DLL" Alias "NewAction" (ByVal lpszKey As String, ByVal lpszData As String) As Long
Private Declare Function DllEnableLogging Lib "VB5STKIT.DLL" Alias "EnableLogging" (ByVal lpszFilename As String) As Long
Private Declare Function DllDisableLogging Lib "VB5STKIT.DLL" Alias "DisableLogging" () As Long

'-----------------------------------------------------------
' Procedimiento: AbortAction
'
' Cancela la acción actual.
'-----------------------------------------------------------
'
Sub AbortAction()
    ShowLoggingError DllAbortAction(), LogErrFatal
End Sub

'-----------------------------------------------------------
' Procedimiento: AddActionNote
'
' Agrega una nota que se escribirá en el archivo de registro
' inmediatamente después de la acción actual
'-----------------------------------------------------------
'
Sub AddActionNote(ByVal strNote As String)
    ShowLoggingError DllAddActionNote(strNote), LogErrOK
End Sub

'-----------------------------------------------------------
' Procedimiento: ChangeActionKey
'
' Cambia la clave de la acción actual.
'-----------------------------------------------------------
'
Sub ChangeActionKey(ByVal strNewKey As String)
    ShowLoggingError DllChangeActionKey(strNewKey), LogErrFatal
End Sub

'-----------------------------------------------------------
' Procedimiento: CommitAction
'
' Marca el final con éxito de la acción actual.
' La acción se registrará en el archivo de registro.
'-----------------------------------------------------------
'
Sub CommitAction()
    ShowLoggingError DllCommitAction(), LogErrFatal
End Sub

'-----------------------------------------------------------
' Procedimiento: DisableLogging
'
' Desactiva el registro para desinstalar aplicaciones. Se podrá llamar
' a todas las funciones de registro y deben ser semánticamente correctas,
' pero no se escribirán datos en el disco.
'-----------------------------------------------------------
'
Sub DisableLogging()
    ShowLoggingError DllDisableLogging(), LogErrFatal
End Sub

'-----------------------------------------------------------
' Procedimiento: EnableLogging
'
' Activa la instalación y desinstalación de aplicaciones al archivo de registro especificado
'-----------------------------------------------------------
'
Sub EnableLogging(ByVal strLogFileName As String)
    ShowLoggingError DllEnableLogging(strLogFileName), LogErrFatal
End Sub

'-----------------------------------------------------------
' Procedimiento: LogError
'
' Registra un error en el archivo de registro. La acción no se cancelará.
'-----------------------------------------------------------
'
Sub LogError(ByVal strErr As String)
    ShowLoggingError DllLogError(ResolveResString(resLOG_ERROR), ResolveResString(resLOG_DURINGACTION), strErr), LogErrFatal
End Sub

'-----------------------------------------------------------
' Procedimiento: LogWarning
'
' Registra una advertencia en el archivo de registro. La acción no se cancelará.
' Las advertencias se diferencian de los errores en que generalmente
' no necesitan la atención del usuario final.
' También, en que el instalador no registra advertencias. Sólo
' registra errores.
'
' Es posible desactivar el registro de advertencias cambiando el valor
' de fLOG_WARNINGS en la sección de declaraciones de este
' módulo.
'-----------------------------------------------------------
'
Sub LogWarning(ByVal strWarning As String)
    If fLOG_WARNINGS Then
        ShowLoggingError DllLogWarning(ResolveResString(resLOG_WARNING), ResolveResString(resLOG_DURINGACTION), strWarning), LogErrFatal
    End If
End Sub

'-----------------------------------------------------------
' Procedimiento: LogNote
'
' Registra una nota en el archivo de registro. No es necesario tener
' una acción actual para ejecutar esta rutina.
'-----------------------------------------------------------
'
Sub LogNote(ByVal strNote As String)
    ShowLoggingError DllLogNote(strNote), LogErrOK
End Sub

'-----------------------------------------------------------
' Procedimiento: NewAction
'
' Marca el inicio de una nueva acción a registrar. Si se llama a
' esta rutina antes de realizar o cancelar la acción actual,
' la acción anterior se situará en una pila.
' Cuando se ha completado o cancelado la nueva acción
' la acción anterior será la actual.
' El registro de errores, advertencias, notas y acciones
' no se imprime hasta que la acción se cancele o se
' complete.
' Es posible apilar varias acciones en una pila FIFO
' llamando a esta rutina sucesivamente.
'-----------------------------------------------------------
'
Sub NewAction(ByVal strKey As String, ByVal strData As String)
    ShowLoggingError DllNewAction(strKey, strData), LogErrFatal
End Sub

Sub ShowLoggingError(ByVal lErr As Long, ByVal lErrSeverity As Long)
    If lErr = LOGERR_SUCCESS Then
        Exit Sub
    End If
    
    Dim strErrMsg As String
    Static fRecursive As Boolean
    
    If fRecursive Then
        'Si estamos llamando recursivamente, es necesario que
        'capturar los errores mientras se intenta escribir en el
        'archivo de registro. Sólo se puede desactivar el registro
        'y cancelar la instalación.
        DisableLogging
        MsgError ResolveResString(resUNEXPECTED), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup frmSetup1, gintRET_FATAL
    End If

    fRecursive = True

    Select Case lErr
        Case LOGERR_OUTOFMEMORY, LOGERR_WRITEERROR, LOGERR_UNEXPECTED, LOGERR_FILENOTFOUND
            strErrMsg = ResolveResString(resUNEXPECTED)
        
        Case LOGERR_INVALIDARGS, LOGERR_EXCEEDEDCAPACITY, LOGERR_NOCURRENTACTION
            'Nota: Estos errores provienen de una personalización inadecuada
            'de este proyecto. Tenga en cuenta que los cambios que haga a estos archivos
            'son válidos y libres de errores.
            'LOGERR_INVALIDARGS -- algún parámetro de una función de registro no es válido o inadecuado
            'LOGERR_EXCEEDEDCAPACITY -- se han sobrepasado los límites de la pila de acciones.
            '   Probablemente no tiene en su código instrucciones CommitAction o
            '   AbortAction.
            'LOGERR_NOCURRENTACTION -- la función de registro ha intentado usar recursos
            '   de una acción actual, pero no había disponibles. Compruebe si falta una
            '   instrucción NewAction.
            strErrMsg = ResolveResString(resUNEXPECTED)
        Case Else
            strErrMsg = ResolveResString(resUNEXPECTED)
        'End Case
    End Select
    
    Dim iRet As Integer
    Dim fAbort As Boolean
    
    fAbort = False
    If lErrSeverity = LogErrOK Then
        ' El usuario puede elegir si desea o no continuar
        iRet = MsgFunc(strErrMsg, MB_OKCANCEL Or MB_ICONEXCLAMATION, gstrTitle)
        If gfNoUserInput Then iRet = IDCANCEL ' imposible continuar en modo silencioso.
        Select Case iRet
            Case IDOK
            Case IDCANCEL
                fAbort = True
            Case Else
                fAbort = True
            'End Case
        End Select
    Else
        ' Fatal
        MsgFunc strErrMsg, MB_OK Or MB_ICONEXCLAMATION, gstrTitle
        fAbort = True
    End If

    If fAbort Then
        ExitSetup frmCopy, gintRET_ABORT
    End If

    fRecursive = False

End Sub

'-----------------------------------------------------------
' Función: fWithinAction
'
' Devuelve TRUE si hay una acción actual
'-----------------------------------------------------------
'
Function fWithinAction() As Boolean
    fWithinAction = fDllWithinAction()
End Function

