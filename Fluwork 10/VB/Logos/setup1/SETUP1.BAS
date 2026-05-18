Attribute VB_Name = "basSetup1"
Option Explicit
Option Compare Text

'
' Constantes globales
'

'Devuelve valores para las funciones del kit de herramientas de instalación
Global Const gintRET_CONT% = 1
Global Const gintRET_CANCEL% = 2
Global Const gintRET_EXIT% = 3
Global Const gintRET_ABORT% = 4
Global Const gintRET_FATAL% = 5
Global Const gintRET_FINISHEDSUCCESS% = 6 'Se usa sólo como parámetro de ExitSetup al final de una instalación correcta

'Niveles de error para GetAppRemovalCmdLine()
Global Const APPREMERR_NONE = 0 'ningún error
Global Const APPREMERR_FATAL = 1 'error fatal
Global Const APPREMERR_NONFATAL = 2 'error no fatal, el usuario eligió anular
Global Const APPREMERR_USERCANCEL = 3 'el usuario eligió cancelar (ningún error)

'Indicador del cuadro de diálogo Ruta de acceso que especifica que se necesita el directorio de origen o de destino
Global Const gstrDIR_SRC$ = "S"
Global Const gstrDIR_DEST$ = "D"

'Comienzo de las líneas de las secciones [Files], [Bootstrap] y [Licenses] de SETUP.LST
Global Const gstrINI_FILE$ = "File"
Global Const gstrINI_REMOTE$ = "Remote"
Public Const gstrINI_LICENSE$ = "License"
'
' Constantes de la línea de comandos
'
Public Const gstrSILENTSWITCH = "s"
Public Const gstrSMSSWITCH = "q"
'
'Definiciones de tipos
'
Type FILEINFO                                               ' Formato de línea del archivo de información de la instalación
    intDiskNum As Integer                                   ' número de disco
    fSplit As Integer                                       ' indicador de división
    strSrcName As String                                    ' nombre del archivo de origen
    strDestName As String                                   ' nombre del archivo de destino
    strDestDir As String                                    ' directorio de destino
    strRegister As String                                   ' información de registro
    fShared As Boolean                                      ' si el archivo está compartido o es privado
    fSystem As Boolean                                      ' si es un archivo de sistema (p.ej., se debe instalar pero no quitar)
    varDate As Variant                                      ' fecha del archivo
    lFileSize As Long                                       ' tamaño del archivo
    sVerInfo As VERINFO                                     ' número de versión del archivo
    strReserved As String                                   ' Reservado. Dejar vacío o se producirá un error.
    strProgramIconTitle As String                           ' Título del icono en el grupo de programas
    strProgramIconCmdLine As String                         ' Línea de comandos para el icono en el grupo de programas
End Type

Type DISKINFO                                               ' Información de la unidad de disco
    lAvail As Long                                          ' Bytes disponibles en la unidad
    lReq As Long                                            ' Bytes necesarios para la instalación
    lMinAlloc As Long                                       ' unidad de asignación mínima
End Type

Type DESTINFO                                               ' guardar directorio de destino para ciertos archivos
    strAppDir As String
    strAUTMGR32 As String
    strRACMGR32 As String
End Type

Type REGINFO                                                ' guardar información de registro para los archivos
    strFilename As String
    strRegister As String
    
    'Lo siguiente sólo se utiliza para registro de servidor remoto
    strNetworkAddress As String
    strNetworkProtocol As String
    intAuthentication As Integer
    fDCOM As Boolean      ' Es True si se usa DCOM; de lo contrario, es False
End Type

'
'Variables globales
'
Global gstrSETMSG As String
Global gfRetVal As Integer                                  'valor de retorno para funciones basadas en formulario
Global gstrAppName As String                                'nombre de la aplicación que se va a instalar
Global gstrTitle As String                                  'nombre de "instalación" de la aplicación que se va a instalar
Public gstrDefGroup As String                               'nombre predeterminado del grupo; de setup.lst
Global gstrDestDir As String                                'directorio de destino para archivos de aplicación
Global gstrAppExe As String                                 'nombre del .EXE de la aplicación que se va a instalar
Public gstrAppToUninstall As String                         'nombre del exe/ocx/dll de la aplicación que se va a desinstalar. Debe ser igual que gstrAppExe en la mayoría de los casos.
Global gstrSrcPath As String                                'ruta de acceso de los archivos de origen
Global gstrSetupInfoFile As String                          'nombre de ruta de acceso del archivo SETUP.LST
Global gstrWinDir As String                                 'directorio de Windows
Global gstrWinSysDir As String                              'directorio windows\system
Global gsDiskSpace() As DISKINFO                            'espacio en disco para las unidades de destino
Global gstrDrivesUsed As String                             'unidades de destino usadas por la instalación
Global glTotalCopied As Long                                'bytes totales copiados hasta ahora
Global gintCurrentDisk As Integer                           'número de disco actual que se está instalando
Global gsDest As DESTINFO                                   'directorios de destino para ciertos archivos
Global gstrAppRemovalLog As String                          'nombre del archivo de registro para quitar aplicaciones
Global gstrAppRemovalEXE As String                          'nombre del ejecutable para quitar aplicaciones
Global gfAppRemovalFilesMoved As Boolean                    'si los archivos para quitar aplicaciones se han movido o no al directorio de la aplicación
Global gfForceUseDefDest As Boolean                         'si se establece a True, no se preguntará al usuario el directorio de destino
Global fMainGroupWasCreated As Boolean                     'Si se ha creado o no una carpeta o grupo principal
Public gfRegDAO As Boolean                                 ' Si se establece a True en el código,
                                                           ' tenemos que agregar al registro
                                                           ' cierta información sobre DAO.


'
'Constantes de formulario/módulo
'

'Posibles acciones de ProgMan
Const mintDDE_ITEMADD% = 1                                  'Indicador AddProgManItem
Const mintDDE_GRPADD% = 2                                   'Indicador AddProgManGroup

'Nombres de archivo especiales
Const mstrFILE_APPREMOVALLOGBASE$ = "ST5UNST"               'Nombre de base del archivo de registro para quitar aplicaciones
Const mstrFILE_APPREMOVALLOGEXT$ = ".LOG"                   'Extensión predeterminada para el archivo de registro para quitar aplicaciones
Const mstrFILE_AUTMGR32 = "AUTMGR32.EXE"
Const mstrFILE_RACMGR32 = "RACMGR32.EXE"
Const mstrFILE_CTL3D32$ = "CTL3D32.DLL"
Const mstrFILE_RICHED32$ = "RICHED32.DLL"

'Nombre del archivo temporal usado para concatenar archivos divididos
Const mstrCONCATFILE$ = "VB5STTMP.CCT"

'Macros de registro del archivo de información de la instalación
Const mstrDLLSELFREGISTER$ = "$(DLLSELFREGISTER)"
Const mstrEXESELFREGISTER$ = "$(EXESELFREGISTER)"
Const mstrTLBREGISTER$ = "$(TLBREGISTER)"
Const mstrREMOTEREGISTER$ = "$(REMOTE)"
Const mstrVBLREGISTER$ = "$(VBLREGISTER)"  ' Error 5-8039

'
'Variables de formulario/módulo
'
Private msRegInfo() As REGINFO                                  'archivos para registrar
Private mlTotalToCopy As Long                                   'bytes totales para copiar
Private mintConcatFile As Integer                               'controlador del archivo de destino para la concatenación
Private mlSpaceForConcat As Long                                'espacio adicional necesario para la concatenación
Private mstrConcatDrive As String                               'unidad que se va a usar para la concatenación
Private mstrVerTmpName As String                                'nombre de archivo temporal para la API VerInstallFile

' Caché de Hkey (usada con fines de registro)
Private Type HKEY_CACHE
    hKey As Long
    strHkey As String
End Type

Private hkeyCache() As HKEY_CACHE

' API de manipulación del registro (32 bits)
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0&
Const ERROR_NO_MORE_ITEMS = 259&

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4


Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszName As String, ByVal cchName As Long) As Long
Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long

Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'-----------------------------------------------------------
' Procedimiento: AddPerAppPath
'
' Agrega al registro del sistema la ruta de acceso completa
'   de la aplicación y la ruta por aplicación (sólo es
'   significativo para Windows 95).
'
' ENTRADA: [strAppExe] - nombre EXE de la aplicación, sin
'       incluir la ruta de acceso
'     [strAppDir] - ruta completa del EXE, sin incluir el
'       nombre de archivo
'     [strAppPath] - ruta por aplicación para esta aplicación
'       (lista separada por signos de punto y coma con los
'       nombres de ruta de directorio)
'       Si es una cadena vacía (""), no hay ninguna ruta por
'       aplicación registrada; se registra el nombre completo
'       de la ruta de acceso del EXE.
'
' SALIDA:
'   Entradas de ejemplo del registro:
'     HKEY_LOCAL_MACHINE\[strPathsBaseKeyName]\MiApl.Exe
'       [Default]=C:\Archivos de programa\MiApl\MiApl.Exe
'       [Path]=C:\Archivos de programa\MiApl;C:\Archivos de programa\MiApl\System
'
'-----------------------------------------------------------
'
Sub AddPerAppPath(ByVal strAppExe As String, ByVal strAppDir As String, ByVal strPerAppPath As String)
    If Not TreatAsWin95() Then
        Exit Sub
    End If
    
    Dim strPathsBaseKeyName As String
    Const strAppPaths$ = "App Paths"
    Const strAppPathKeyName = "Path"
    Dim fOk As Boolean
    Dim hKey As Long
    
    AddDirSep strAppDir
    
    ' Crea la nueva clave, cuyo nombre se basa en el nombre de la aplicación
   If Not RegCreateKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), strAppPaths & gstrSEP_DIR & strAppExe, hKey) Then
        GoTo Err
    End If
    
    fOk = True
    
    ' El valor predeterminado indica el nombre de ruta completo del EXE
    fOk = fOk And RegSetStringValue(hKey, "", strAppDir & strAppExe)
    
    ' El valor de [Path] indica la ruta de acceso por aplicación
    If strPerAppPath <> "" Then
        fOk = fOk And RegSetStringValue(hKey, strAppPathKeyName, strPerAppPath)
    End If
    
    If Not fOk Then
        GoTo Err
    End If
    
    RegCloseKey hKey
    
    Exit Sub
    
Err:
    MsgError ResolveResString(resERR_REG), vbExclamation Or vbOKOnly, gstrTitle
    '
    ' Si estamos realizando una instalación en SMS, no podemos continuar.
    '
    If gfSMS Then
        ExitSetup frmSetup1, gintRET_FATAL
    End If
End Sub

'-----------------------------------------------------------
' Función: AddQuotesToFN
'
' Dado un nombre de ruta de acceso (directorio y/o nombre
'   de archivo), devuelve dicho nombre de ruta entre dobles
'   comillas si la ruta contiene espacios en blanco o
'   comas. Esto es necesario para configurar un icono
'   correctamente, ya que de lo contrario dichas rutas
'   se interpretarían como un nombre de ruta de acceso
'   con argumentos.
'-----------------------------------------------------------
'
Function AddQuotesToFN(ByVal strFilename) As String
    If InStr(strFilename, " ") Or InStr(strFilename, ",") Then
        AddQuotesToFN = """" & strFilename & """"
    Else
        AddQuotesToFN = strFilename
    End If
End Function

'-----------------------------------------------------------
' Procedimiento: CalcDiskSpace
'
' Calcula el espacio en disco necesario para instalar los
' archivos enumerados en la sección especificada del archivo
' de información de configuración (SETUP.LST)
'-----------------------------------------------------------
'
Sub CalcDiskSpace(ByVal strSection As String)
    Static fSplitFile As Integer
    Static lDestFileSpace As Long

    Dim intIdx As Integer
    Dim intDrvIdx As Integer
    Dim sFile As FILEINFO
    Dim strDrive As String
    Dim lThisFileSpace As Long

    intIdx = 1

    On Error GoTo CalcDSError

    '
    'Para cada archivo de la sección especificada, lee información del archivo de
    'información de configuración
    '
    Do While ReadSetupFileLine(strSection, intIdx, sFile) = True
        '
        'si el archivo no está dividido o si ésta es la primera sección de un
          'archivo dividido
        '
        If sFile.strDestDir <> gstrNULL Then
            fSplitFile = sFile.fSplit

            '
            'Obtiene la unidad de prueba empleada para este archivo. Si es el primer
                'archivo que usa la unidad para un destino, agrega la unidad a la 'tabla'
                'de unidades usadas, asigna un elemento de matriz para almacenar la
                'información de unidades y obtiene el espacio disponible en disco y la
                'unidad mínima de asignación
            '
            strDrive = Left$(sFile.strDestDir, 1)
        
            intDrvIdx = InStr(gstrDrivesUsed, strDrive)
            If intDrvIdx = 0 Then
                gstrDrivesUsed = gstrDrivesUsed & strDrive
                intDrvIdx = Len(gstrDrivesUsed)

                ReDim Preserve gsDiskSpace(intDrvIdx)
                gsDiskSpace(intDrvIdx).lAvail = GetDiskSpaceFree(strDrive)

                gsDiskSpace(intDrvIdx).lMinAlloc = GetDrivesAllocUnit(strDrive)
            End If

            '
            'Calcula el tamaño del destino final (tamaño de archivo + asignación mínima
                'para la unidad)
            '
            lThisFileSpace = CalcFinalSize(sFile.lFileSize, strDrive)
            mlTotalToCopy = mlTotalToCopy + lThisFileSpace

            '
            'Si el archivo ya existe, si lo copiamos lo reemplazaremos. Por eso,
            'obtenemos el tamaño del archivo de destino existente de forma que lo
                'podamos restar de la cantidad necesaria posteriormente.
            '
            If FileExists(sFile.strDestDir & sFile.strDestName) Then
                lDestFileSpace = FileLen(sFile.strDestDir & sFile.strDestName)
            Else
                lDestFileSpace = 0
            End If
        End If

        '
        'Si el archivo no está dividido o si es la última sección de un archivo dividido
        '
        If sFile.fSplit = False Then
            '
            'Si es la última sección de un archivo dividido, es el *mayor* archivo
            'dividido; establece el espacio adicional necesario para la concatenación
                'a este tamaño
            '
            If fSplitFile = True And lThisFileSpace > mlSpaceForConcat Then
                mlSpaceForConcat = lThisFileSpace
            End If

            '
            'Resta el tamaño del archivo de destino existente, si es aplicable, y después
            'acumula el espacio necesario
            '
            lThisFileSpace = lThisFileSpace - lDestFileSpace
            If lThisFileSpace < 0 Then
                lThisFileSpace = 0
            End If

            gsDiskSpace(intDrvIdx).lReq = gsDiskSpace(intDrvIdx).lReq + lThisFileSpace
        End If

        intIdx = intIdx + 1
    Loop

    Exit Sub

CalcDSError:
    MsgError Error$ & LS$ & ResolveResString(resCALCSPACE), MB_ICONSTOP, gstrSETMSG
    ExitSetup frmMessage, gintRET_FATAL
End Sub

'-----------------------------------------------------------
' Procedimiento: CalcFinalSize
'
' Calcula el espacio necesario para un archivo del tamaño
' especificado en la ruta de destino indicada. Esto incluye
' el tamaño de archivo, así como un relleno, para asegurar
' que el tamaño final es múltiplo de la unidad mínima de
' asignación para la unidad de destino.
'-----------------------------------------------------------
'
Function CalcFinalSize(lBaseFileSize As Long, strDestPath As String) As Long
    Dim lMinAlloc As Long
    Dim intPadSize As Long

    lMinAlloc = gsDiskSpace(InStr(gstrDrivesUsed, Left$(strDestPath, 1))).lMinAlloc
    intPadSize = lMinAlloc - (lBaseFileSize Mod lMinAlloc)
    If intPadSize = lMinAlloc Then
        intPadSize = 0
    End If

    CalcFinalSize = lBaseFileSize + intPadSize
End Function

'-----------------------------------------------------------
' Procedimiento: CenterForm
'
' Centra el formulario pasado, justo encima del centro de
' la pantalla.
'-----------------------------------------------------------
'
Sub CenterForm(frm As Form)
    SetMousePtr gintMOUSE_HOURGLASS

    frm.Top = (Screen.Height * 0.85) \ 2 - frm.Height \ 2
    frm.Left = Screen.Width \ 2 - frm.Width \ 2

    SetMousePtr gintMOUSE_DEFAULT
End Sub

'-----------------------------------------------------------
' Función: CheckDiskSpace
'
' Lee de la matriz de espacio necesario que se genera al
' llamar a la función 'CalcDiskSpace' y determina si hay
' suficiente espacio libre en todas las unidades empleadas
' para la instalación.
'
' Retorno: True si hay suficiente espacio; de lo contrario,
' False
'-----------------------------------------------------------
'
Function CheckDiskSpace() As Integer
    Static fDontAskOnSpaceErr As Integer

    Dim intIdx As Integer
    Dim intTmpDrvIdx As Integer
    Dim lDiskSpaceLeft As Long
    Dim lMostSpaceLeft As Long
                                             
    '
    'De forma predeterminada es True (suficiente espacio en todas las unidades)
    '
    CheckDiskSpace = True

    '
    'Para cada unidad que es destino de uno o más archivos, compara el espacio
    'disponible con el espacio necesario.
    '
    For intIdx = 1 To Len(gstrDrivesUsed)
        lDiskSpaceLeft = gsDiskSpace(intIdx).lAvail - gsDiskSpace(intIdx).lReq
        If lDiskSpaceLeft < 0 Then
            GoSub CheckDSAskSpace
        Else
            '
            'Si no se encuentra ninguna unidad "TMP", o si la unidad "TMP" no
                'estaba lista, guarda el índice de la unidad y la cantidad de
                'espacio de la unidad que tendrá más espacio libre. Si no se
                'encuentra ninguna unidad "TMP" en InitDiskInfo(), esta unidad se
                'usará como unidad temporal para concatenar archivos divididos
            '
            If mstrConcatDrive = gstrNULL Then
                If lDiskSpaceLeft > lMostSpaceLeft Then
                    lMostSpaceLeft = lDiskSpaceLeft
                    intTmpDrvIdx = intIdx
                End If
            Else
                '
                'Se especificó la unidad "TMP"; es la que usaremos
                '
                If Left$(mstrConcatDrive, 1) = Mid$(gstrDrivesUsed, intIdx, 1) Then
                    intTmpDrvIdx = intIdx
                End If
            End If
        End If
    Next

    '
    'Si al menos se especificó una unidad como destino (si había al menos una
    'llamada CalcDiskSpace en Form_Load de SETUP1.FRM), restamos el espacio
    'adicional necesario para la concatenación, de:
    '   La unidad "TMP" si está disponible  - O BIEN -
    '   La unidad donde queda más espacio
    '
    If intTmpDrvIdx > 0 Then
        gsDiskSpace(intTmpDrvIdx).lReq = gsDiskSpace(intTmpDrvIdx).lReq + mlSpaceForConcat
        If gsDiskSpace(intTmpDrvIdx).lAvail < gsDiskSpace(intTmpDrvIdx).lReq Then
            GoSub CheckDSAskSpace
        End If

        '
        'Si se encontró una unidad "TMP", la usamos; de lo contrario, usamos la
        'unidad que tenga más espacio libre
        '
        If mstrConcatDrive = gstrNULL Then
            mstrConcatDrive = Mid$(gstrDrivesUsed, intTmpDrvIdx, 1) & gstrCOLON & gstrSEP_DIR
            AddDirSep mstrConcatDrive
        End If
    End If

    Exit Function

CheckDSAskSpace:
    '
    'Si no se ha preguntado antes al usuario en caso de que no haya suficiente
    'espacio libre, muestra la tabla de espacio en las unidades y se le permite
    '(básicamente) anular, reintentar o ignorar.
    '
    If fDontAskOnSpaceErr = False Then
        If gfNoUserInput Then
            If gfSilent = True Then
                LogSilentMsg ResolveResString(resLBLNOSPACE)
            End If
            If gfSMS = True Then
                LogSMSMsg ResolveResString(resLBLNOSPACE)
            End If
            ExitSetup frmSetup1, gintRET_FATAL
        Else
            frmDskSpace.Show 1
        End If
        
        If gfRetVal <> gintRET_CONT Then
            CheckDiskSpace = False
            Exit Function
        Else
            fDontAskOnSpaceErr = True
        End If
    End If

    Return
End Function

'-----------------------------------------------------------
' Función: CheckDrive
'
' Comprueba si se puede leer de la unidad especificada. En
' el caso de una unidad que contiene un medio extraíble,
' esto significaría que en la unidad había un medio con
' formato y que la puerta de la unidad estaba cerrada.
'
' ENTRADA: [strDrive] - unidad que se va a comprobar
'     [strCaption] - título si la unidad no está lista
'
' Retorno: True si la unidad está lista; de lo contrario, False
'-----------------------------------------------------------
'
Function CheckDrive(ByVal strDrive As String, ByVal strCaption As String) As Integer
    Dim strDir As String
    Dim strMsg As String
    Dim fIsUNC As Boolean

    On Error Resume Next

    SetMousePtr gintMOUSE_HOURGLASS

    Do
        Err = 0
        fIsUNC = False
        '
        'Intenta leer el directorio actual de la unidad especificada.
        'Si se produce un error, suponemos que la unidad no está lista.
        '
        If IsUNCName(strDrive) Then
            fIsUNC = True
            strDir = Dir$(GetUNCShareName(strDrive))
        Else
            strDir = Dir$(Left$(strDrive, 2))
        End If

        If Err > 0 Then
            If fIsUNC Then
                strMsg = Error$ & LS$ & ResolveResString(resCANTREADUNC, "|1", strDrive) & LS$ & ResolveResString(resCHECKUNC)
            Else
                strMsg = Error$ & LS$ & ResolveResString(resDRVREAD) & strDrive & LS$ & ResolveResString(resDRVCHK)
            End If
            If MsgError(strMsg, MB_ICONEXCLAMATION Or MB_RETRYCANCEL, strCaption) = IDCANCEL Then
                CheckDrive = False
                Err = 0
            End If
        Else
            CheckDrive = True
        End If
        
        If Err And gfNoUserInput = True Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    Loop While Err

    SetMousePtr gintMOUSE_DEFAULT
End Function

'-----------------------------------------------------------
' Función: CheckOverwritePrivateFile
'
' Comprueba si un archivo privado que estamos a punto de
' instalar ya existe en el directorio de destino. Si existe,
' habrá problemas en caso de que el usuario intente quitar
' alguna aplicación; por tanto, avise al usuario y sugiérale
' que seleccione otro directorio de destino diferente.
'
' ENTRADA: [strFN] - Ruta completa del archivo privado que
'               se va a instalar.
'
'-----------------------------------------------------------
'
Sub CheckOverwritePrivateFile(ByVal strFN As String)
    Static fIgnoreOverwrite As Boolean
    
    If fIgnoreOverwrite Then
        'Si el usuario elige una vez ignorar esta advertencia,
        'no la mostraremos de nuevo.
        Exit Sub
    End If
    
    If FileExists(strFN) Then
        Do
            Select Case MsgError(ResolveResString(resOVERWRITEPRIVATE) & LS$ & ResolveResString(resCANCELSETUP), vbYesNo Or vbDefaultButton1 Or vbExclamation, gstrTitle)
                Case vbYes
                    'El usuario eligió cancelar.  (Es lo mejor.)
                    gfDontLogSMS = True  ' No registra este mensaje en SMS porque ya registramos el anterior y sólo podemos usar 255 caracteres.
                    MsgError ResolveResString(resCHOOSENEWDEST), vbOKOnly, gstrTitle
                    ExitSetup frmCopy, gintRET_FATAL
                Case Else
                    'Un  nivel más de advertencia para hacerle saber que en este
                    '  momento recomendamos cancelar la instalación
                    Select Case MsgError(ResolveResString(resOVERWRITEPRIVATE2) & LS$ & ResolveResString(resVERIFYCONTINUE), vbYesNo Or vbDefaultButton2 Or vbExclamation, gstrTitle)
                        Case vbNo
                            'El usuario eligió "no, no continuar"
                            'Repite la advertencia de primer nivel
                        Case Else
                            'Decidió continuar de todas maneras
                            fIgnoreOverwrite = True
                            Exit Do
                        'End Case
                    End Select
                'End Case
            End Select
        Loop
    End If
End Sub

'-----------------------------------------------------------
' Función: ConcatSplitFile
'
' Lee y anexa el archivo de origen pasado al archivo de
' destino abierto anteriormente especificado por
' mintConcatFile. Se debe abrir mintConcatFile llamando a
' OpenConcatFile() antes de llamar a esta función.
'
' ENTRADA: [strSrcName] - Archivo de origen que se va a
'          anexar al destino
'
' Retorno: True si la copia se realizó correctamente; IDIGNORE
'          si el usuario elige ignorar un error de copia del
'          que se ha informado
'-----------------------------------------------------------
'
Function ConcatSplitFile(ByVal strSrcName As String) As Integer
    Const lMAXCOPYBUF& = 64512
    Const lMINCOPYBUFSIZE& = 4096
    Const intOPEN% = 1
    Const intGET% = 2
    Const intPUT% = 3
    Const intMEMFAIL% = 4

    Dim intSrcFile As Integer
    Dim intStatus As Integer
    Dim lBytesLeftToWrite As Long
    Dim lBytesThisTime As Long
    Dim byteFileBuf() As Byte 'Debe ser Byte, no String, por lo que no se realiza ninguna conversión Unicode
    Dim strMsg As String

    On Error GoTo CSFError
    
    '
    'Se asegura de que el archivo de origen especificado está disponible
    '
    If DetectFile(strSrcName) = IDIGNORE Then
        ConcatSplitFile = IDIGNORE
        Exit Function
    End If

    lBytesLeftToWrite = FileLen(strSrcName)

    '
    'Para informar de errores, indicador de que estamos intentando abrir el archivo ahora
    '
    intStatus = intOPEN

    '
    'Abre ahora el archivo de origen para lectura
    '
    intSrcFile = FreeFile
    Open strSrcName For Binary Access Read As intSrcFile

    '
    'Inicialmente, intentaremos copiar lMAXCOPYBUF bytes cada vez. Si falla nuestro
    'intento de asignar un búfer de copia (Space$(...)), la lógica de control de errores
    'reducirá a la mitad el tamaño del búfer y se intentará de nuevo otra asignación.
    '
    lBytesThisTime = lMAXCOPYBUF
    ReDim byteFileBuf(1 To lBytesThisTime) As Byte

    While lBytesLeftToWrite <> 0
        '
        'mientras no se ha leído el archivo de origen, si el número de bytes
          'restante es mayor que el tamaño del búfer, reduce el tamaño del búfer
        '
        If lBytesThisTime > lBytesLeftToWrite Then
            lBytesThisTime = lBytesLeftToWrite
            ReDim byteFileBuf(1 To lBytesThisTime) As Byte
        End If
        
        '
        'Establece el estado de la operación y lo obtiene (Get) del archivo de origen
          'y lo pone (Put) en el archivo de destino
        '
        intStatus = intGET
        Get intSrcFile, , byteFileBuf

        intStatus = intPUT
        Put mintConcatFile, , byteFileBuf

        lBytesLeftToWrite = lBytesLeftToWrite - lBytesThisTime
    Wend

    ConcatSplitFile = True
    GoTo CSFCleanup

CSFError:
    If Err = 14 Then    'No hay espacio de cadena
        lBytesThisTime = lBytesThisTime \ 2
        If lBytesThisTime >= lMINCOPYBUFSIZE Then
            Resume
        Else
            intStatus = intMEMFAIL
        End If
    End If

    strMsg = LF$ & strSrcName

    Select Case intStatus
        Case intOPEN
            strMsg = ResolveResString(resCANTOPEN) & strMsg
        Case intGET
            strMsg = ResolveResString(resCANTREAD) & strMsg
        Case intPUT
            strMsg = ResolveResString(resCANTWRITE) & strMsg & LS$ & ResolveResString(resCHKSPACE)
        Case intMEMFAIL
            strMsg = ResolveResString(resOUTOFMEMORY) & strMsg
        'End Case
    End Select

    Select Case MsgError(Error$ & LS$ & strMsg, MB_ABORTRETRYIGNORE Or MB_ICONEXCLAMATION Or IIf(gfNoUserInput, vbDefaultButton1, MB_DEFBUTTON2), gstrSETMSG)
        Case IDABORT
            ExitSetup frmCopy, gintRET_ABORT
        Case IDIGNORE
            ConcatSplitFile = IDIGNORE
        Case IDRETRY
            Resume
        'End Case
    End Select

CSFCleanup:
    Close intSrcFile
    Err = 0
    Exit Function
End Function

'-----------------------------------------------------------
' Función: CopyFile
'
' Usa la API VerInstallFile de Windows para copiar un archivo
' de la ubicación/nombre de origen especificado a la
' ubicación/nombre de destino. Los archivos divididos deben
' combinarse mediante las rutinas de archivo '...Concat...'
' antes de llamar a esta función.
' Si el archivo se actualiza correctamente y es un archivo
' compartido (fShared = True), entonces se actualiza el
' número de referencias del archivo (sólo para 32 bits).
'
' ENTRADA: [strSrcDir] - directorio donde se encuentra el
'       archivo de origen
'     [strDestDir] - directorio de destino del archivo
'     [strSrcName] - nombre del archivo de origen
'     [strDestName] - nombre del archivo de destino
'
' CONDICIÓN PREVIA: Ya se debe haber llamado a NewAction()
'                   para esta copia del archivo (de tipo
'               gstrKEY_SHAREDFILE o gstrKEY_PRIVATE; en
'               CopySection se muestra un ejemplo de cómo
'               funciona esto). Vea NewAction() y las
'                   funciones relacionadas en LOGGING.BAS para
'               ver comentarios acerca de cómo usar la
'               función de registro.
'               Tanto a CommitAction() como a AbortAction() se
'               les debe llamar mediante este procedimiento,
'               no mediante quien hace la llamada.
'
' Retorno: True si la copia se realizó correctamente; de lo
' contrario, False
'
' CONDICIÓN POSTERIOR: La acción actual se confirmará o
'                anulará.
'-----------------------------------------------------------
'
Function CopyFile(ByVal strSrcDir As String, ByVal strDestDir As String, ByVal strSrcName As String, ByVal strDestName As String, ByVal fShared As Boolean, ByVal fSystem As Boolean) As Boolean
    Const intUNKNOWN% = 0
    Const intCOPIED% = 1
    Const intNOCOPY% = 2
    Const intFILEUPTODATE% = 3

    '
    'Indicadores de VerInstallFile()
    '
    Const VIFF_FORCEINSTALL% = &H1
    Const VIF_TEMPFILE& = &H1
    Const VIF_SRCOLD& = &H4
    Const VIF_DIFFLANG& = &H8
    Const VIF_DIFFCODEPG& = &H10
    Const VIF_DIFFTYPE& = &H20
    Const VIF_WRITEPROT& = &H40
    Const VIF_FILEINUSE& = &H80
    Const VIF_OUTOFSPACE& = &H100
    Const VIF_ACCESSVIOLATION& = &H200
    Const VIF_SHARINGVIOLATION = &H400
    Const VIF_CANNOTCREATE = &H800
    Const VIF_CANNOTDELETE = &H1000
    Const VIF_CANNOTRENAME = &H2000
    Const VIF_OUTOFMEMORY = &H8000&
    Const VIF_CANNOTREADSRC = &H10000
    Const VIF_CANNOTREADDST = &H20000
    Const VIF_BUFFTOOSMALL = &H40000

    Static fIgnoreWarn As Integer             '¿se ha avisado al usuario sobre ignorar el error?

    Dim strMsg As String
    Dim lRC As Long
    Dim lpTmpNameLen As Long
    Dim intFlags As Integer
    Dim intRESULT As Integer
    Dim fFileAlreadyExisted

    On Error Resume Next

    CopyFile = False

    '
    'Se asegura de que el archivo de origen está disponible para copiar
    '
    If DetectFile(strSrcDir & strSrcName) = IDIGNORE Then
        AbortAction
        Exit Function
    End If
    
    '
    ' Se asegura de que la ruta de acceso de destino (incluyendo ruta, nombre de archivo,
    ' argumentos de la línea de comandos, etc.) no es mayor que el máximo permitido.
    '
    If Not fCheckFNLength(strDestDir & strDestName) Then
        AbortAction
        strMsg = ResolveResString(resCANTCOPYPATHTOOLONG) & LS$ & ResolveResString(resCHOOSENEWDEST) & LS$ & strDestDir & strDestName
        Call MsgError(strMsg, vbOKOnly, gstrSETMSG)
        ExitSetup frmCopy, gintRET_FATAL
        Exit Function
    End If
    '
    'Crea el directorio de destino; pide al usuario que reintente si hay algún error
    '
    If Not MakePath(strDestDir) Then
        AbortAction ' Anula la copia del archivo
        Exit Function
    End If

    '
    'Se asegura de que tenemos el LFN (nombre largo de archivo) del directorio de destino
    '
    strDestDir = GetLongPathName(strDestDir)
    
    '
    'Configuración para la llamada a VerInstallFile
    '
    lpTmpNameLen = gintMAX_SIZE
    mstrVerTmpName = String$(lpTmpNameLen, 0)
    intFlags = 0
    fFileAlreadyExisted = FileExists(strDestDir & strDestName)

    intRESULT = intUNKNOWN

    Do While intRESULT = intUNKNOWN
        'VerInstallFile bajo Windows 95 no admite nombres largos
        '  de archivo, por lo que tenemos que indicarle las
        '  versiones cortas (sólo para 32 bits).
        Dim strShortSrcName As String
        Dim strShortDestName As String
        Dim strShortSrcDir As String
        Dim strShortDestDir As String
        
        strShortSrcName = strSrcName
        strShortSrcDir = strSrcDir
        strShortDestName = strDestName
        strShortDestDir = strDestDir
        If Not FileExists(strDestDir & strDestName) Then
            'Si el archivo de destino no existe todavía,
            '  creamos uno ficticio con el nombre (largo)
            '  de archivo correcto, de forma que podamos
                '  obtener su nombre corto para VerInstallFile.
            Open strDestDir & strDestName For Output Access Write As #1
            Close #1
        End If
    
        On Error GoTo UnexpectedErr
        If Not IsWindowsNT() Then
            'Esta conversión no es necesaria bajo para Windows NT
            strShortSrcDir = GetShortPathName(strSrcDir)
            strShortSrcName = GetFileName(GetShortPathName(strSrcDir & strSrcName))
            strShortDestDir = GetShortPathName(strDestDir)
            strShortDestName = GetFileName(GetShortPathName(strDestDir & strDestName))
        End If
        On Error Resume Next
            
        lRC = VerInstallFile(intFlags, strShortSrcName, strShortDestName, strShortSrcDir, strShortDestDir, 0&, mstrVerTmpName, lpTmpNameLen)
        If Err <> 0 Then
            '
            'Si no se encuentra las DLL de versión o de expansión de archivo, se anula la instalación
            '
            ExitSetup frmCopy, gintRET_FATAL
        End If

        If lRC = 0 Then
            '
            'El archivo se instaló correctamente; se incrementa el número de referencia si es necesario
            '
            
            'Una cosa más para nombres largos de archivo: VerInstallFile puede
                'haber cambiado el nombre del archivo a su versión corta si terminamos
                'la copia. Por tanto, simplemente cambiamos de nuevo su nombre al
            'que debería tener.
            Name strDestDir & strShortDestName As strDestDir & strDestName
            intRESULT = intCOPIED
        ElseIf lRC And VIF_SRCOLD Then
            '
            'El archivo de origen era más antiguo, por lo que no se copió; se
                'usará la versión existente del archivo. Se incrementa el número
            'de referencia si es necesario.
            '
            intRESULT = intFILEUPTODATE
        ElseIf lRC And (VIF_DIFFLANG Or VIF_DIFFCODEPG Or VIF_DIFFTYPE) Then
            '
            'Reintentamos y forzamos la instalación para estos casos. Puede modificar
                'este código para preguntar al usuario qué desea hacer.
            '
            intFlags = VIFF_FORCEINSTALL
        ElseIf lRC And VIF_WRITEPROT Then
            strMsg = ResolveResString(resWRITEPROT)
            GoSub CFMsg
        ElseIf lRC And VIF_FILEINUSE Then
            strMsg = ResolveResString(resINUSE)
            GoSub CFMsg
        ElseIf lRC And VIF_OUTOFSPACE Then
            strMsg = ResolveResString(resOUTOFSPACE) & Left$(strDestDir, 2)
            GoSub CFMsg
        ElseIf lRC And VIF_ACCESSVIOLATION Then
            strMsg = ResolveResString(resACCESSVIOLATION)
            GoSub CFMsg
        ElseIf lRC And VIF_SHARINGVIOLATION Then
            strMsg = ResolveResString(resSHARINGVIOLATION)
            GoSub CFMsg
        ElseIf lRC And VIF_OUTOFMEMORY Then
            strMsg = ResolveResString(resOUTOFMEMORY)
            GoSub CFMsg
        Else
            '
            ' En estos casos, en general se informa del error y no se instala el archivo
            ' a menos que sea una instalación en SMS, en cuyo caso anulamos la operación.
            '
            If lRC And VIF_CANNOTCREATE Then
                strMsg = ResolveResString(resCANNOTCREATE)
            ElseIf lRC And VIF_CANNOTDELETE Then
                strMsg = ResolveResString(resCANNOTDELETE)
            ElseIf lRC And VIF_CANNOTRENAME Then
                strMsg = ResolveResString(resCANNOTRENAME)
            ElseIf lRC And VIF_CANNOTREADSRC Then
                strMsg = ResolveResString(resCANNOTREADSRC)
            ElseIf lRC And VIF_CANNOTREADDST Then
                strMsg = ResolveResString(resCANNOTREADDST)
            ElseIf lRC And VIF_BUFFTOOSMALL Then
                strMsg = ResolveResString(resBUFFTOOSMALL)
            End If

            strMsg = strMsg & ResolveResString(resNOINSTALL)
            MsgError strMsg, MB_OK Or MB_ICONEXCLAMATION, gstrTitle
            If gfSMS Then
                ExitSetup frmSetup1, gintRET_FATAL
            End If
            intRESULT = intNOCOPY
        End If
    Loop

    '
    'Si quedó un archivo temporal de VerInstallFile, se quita
    '
    If lRC And VIF_TEMPFILE Then
        Kill mstrVerTmpName
    End If

    'Anula o confirma la acción actual y se cuentan las referencias
    Select Case intRESULT
        Case intNOCOPY
            AbortAction
        Case intCOPIED
            DecideIncrementRefCount strDestDir & strDestName, fShared, fSystem, fFileAlreadyExisted
            AddActionNote ResolveResString(resLOG_FILECOPIED)
            CommitAction
            CopyFile = True
        Case intFILEUPTODATE
            DecideIncrementRefCount strDestDir & strDestName, fShared, fSystem, fFileAlreadyExisted
            AddActionNote ResolveResString(resLOG_FILEUPTODATE)
            CommitAction
            CopyFile = True
        Case Else
            AbortAction ' Defensivo: no se debería llegar aquí
        'End Case
    End Select

    Exit Function

UnexpectedErr:
    MsgError Error$ & LS$ & ResolveResString(resUNEXPECTED), vbOKOnly Or vbExclamation, gstrTitle
    ExitSetup frmCopy, gintRET_FATAL
    
CFMsg: '(Subrutina)
    Dim intMsgRet As Integer
    strMsg = strDestDir & strDestName & LS$ & strMsg
    intMsgRet = MsgError(strMsg, MB_ABORTRETRYIGNORE Or MB_ICONEXCLAMATION Or MB_DEFBUTTON2, gstrTitle)
    If gfNoUserInput Then intMsgRet = IDABORT
    Select Case intMsgRet
        Case IDABORT
            ExitSetup frmCopy, gintRET_ABORT
        Case IDIGNORE
            If fIgnoreWarn = True Then
                intRESULT = intNOCOPY
            Else
                fIgnoreWarn = True
                strMsg = strMsg & LS$ & ResolveResString(resWARNIGNORE)
                If MsgError(strMsg, MB_YESNO Or MB_ICONQUESTION Or MB_DEFBUTTON2, gstrTitle) = IDYES Then
                    intRESULT = intNOCOPY
                Else
                    'Will retry
                End If
            End If
        'End Case
    End Select

    Return
End Function

'-----------------------------------------------------------
' Procedimiento: CopySection
'
' Intenta copiar los archivos necesarios desde la sección
' especificada del archivo de información de la instalación
' (SETUP.LST)
'
' ENTRADA: [strSection] - nombre de la sección de la que
' se van a copiar los archivos
'
'-----------------------------------------------------------
'
Sub CopySection(ByVal strSection As String)
    Dim intIdx As Integer
    Dim fSplit As Integer
    Dim fSrcVer As Integer
    Dim sFile As FILEINFO
    Dim strLastFile As String
    Dim intRC As Integer
    Dim lThisFileSize As Long
    Dim strSrcDir As String
    Dim strDestDir As String
    Dim strSrcName As String
    Dim strDestName As String
    Dim strRegister As String
    Dim sSrcVerInfo As VERINFO
    Dim sDestVerInfo As VERINFO
    Dim fFileWasUpToDate As Boolean
    Dim strMultDirBaseName As String
    Dim strMsg As String
    Dim strDetectPath As String
    Dim fRemoteReg As Boolean

    On Error Resume Next

    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    intIdx = 1

    '
    'Para cada archivo de la sección especificada, lee información del
    'archivo de información de la instalación
    '
    Do While ReadSetupFileLine(strSection, intIdx, sFile) = True
        fFileWasUpToDate = False
        
        '
        'Si el último resultado era IGNORAR y esto es un fragmento de un
          'archivo dividido, no es necesario procesar este fragmento
        '
        If intRC = IDIGNORE And sFile.strDestName = strDestName Then
            GoTo CSContinue
        End If

        intRC = 0

        '
        ' Si se requiere un nuevo disco, o si por alguna razón no encontramos
          ' la ruta de acceso de origen (por ejemplo, el usuario quitó el disco
          ' de instalación), pedimos el siguiente disco. La función PromptForNextDisk
          ' no preguntará nada al usuario a menos que determine que la unidad
        ' de origen contiene un medio extraíble o es una conexión de red. Además,
          ' no se pregunta si es una instalación desatendida. Fallará más adelante
          ' en una instalación desatendida cuando no pueda encontrar el archivo.
        '
        If gfNoUserInput = False And (sFile.intDiskNum <> gintCurrentDisk Or DirExists(gstrSrcPath) = False) Then
            PromptForNextDisk sFile.intDiskNum, sFile.strSrcName
        End If

        strSrcName = sFile.strSrcName
        '
        ' El archivo podría existir en el directorio principal de origen o en un
        ' subdirectorio llamado DISK1, DISK2, etc. Se establece la ruta de acceso
        ' apropiada. Si no está en ningún lugar, se trata de un error y se
        ' resolverá más adelante.
        '
        If FileExists(gstrSrcPath & strSrcName) = True Then
            strSrcDir = gstrSrcPath
        ElseIf FileExists(gstrSrcPath & ".." & gstrSEP_DIR & strMultDirBaseName & Format(sFile.intDiskNum) & gstrSEP_DIR & strSrcName) = True Then
            strSrcDir = ResolveDir(gstrSrcPath & ".." & gstrSEP_DIR & strMultDirBaseName & Format(sFile.intDiskNum) & gstrSEP_DIR, False, False)
            gstrSrcPath = strSrcDir
        Else
            '
            ' No se encuentra el archivo.
            '
            If DirExists(gstrSrcPath & strMultDirBaseName & Format(sFile.intDiskNum)) = True Then
                strDetectPath = gstrSrcPath & strMultDirBaseName & Format(sFile.intDiskNum)
            Else
                strDetectPath = gstrSrcPath
            End If
            strMsg = ResolveResString(resCOMMON_CANTFINDSRCFILE, "|1", strDetectPath & gstrSEP_DIR & strSrcName)
            MsgError strMsg, vbExclamation Or vbOKOnly, gstrTitle
            ExitSetup frmCopy, gintRET_FATAL
        End If

        '
        'Si el archivo no está dividido o si es la primera sección de un archivo dividido
        '
        If sFile.strDestDir <> gstrNULL Then
            fSplit = sFile.fSplit

            strDestDir = sFile.strDestDir
            strDestName = sFile.strDestName
            
            'Necesitamos seguir adelante y crear el directorio de destino; de lo contrario,
            'GetLongPathName() puede fallar
            If Not MakePath(strDestDir) Then
                intRC = IDIGNORE
            End If
            
            If intRC <> IDIGNORE Then
                Err = 0
                strDestDir = GetLongPathName(strDestDir)

                frmCopy.lblDestFile.Caption = strDestDir & sFile.strDestName
                frmCopy.lblDestFile.Refresh

                If UCase(strDestName) = gstrFILE_AXDIST Then
                    '
                    ' AXDIST.EXE se instaló temporalmente. Lo eliminaremos al final
                    ' de la instalación. Se establece gfAXDist = True, por lo que
                    ' tenemos que eliminarlo más adelante.
                    '
                    NewAction gstrKEY_TEMPFILE, """" & strDestDir & strDestName & """"
                    gfAXDist = True
                    gstrAXDISTInstallPath = strDestDir & strDestName
                ElseIf UCase(strDestName) = gstrFILE_WINT351 Then
                    '
                    ' WINt351.EXE se instaló temporalmente. Lo eliminaremos al final
                    ' de la instalación. Se establece WINt351 = True, por lo que
                    ' tenemos que eliminarlo más adelante. (Observe que este archivo
                    ' sólo se instala si el destino es nt3.51. Esto se trata con lo
                    ' siguiente en esta misma rutina. Rick Andrews)
                    '
                    NewAction gstrKEY_TEMPFILE, """" & strDestDir & strDestName & """"
                    gfWINt351 = True
                    gstrWINt351InstallPath = strDestDir & strDestName
                ElseIf sFile.fShared Then
                    NewAction gstrKEY_SHAREDFILE, """" & strDestDir & strDestName & """"
                ElseIf sFile.fSystem Then
                    NewAction gstrKEY_SYSTEMFILE, """" & strDestDir & strDestName & """"
                Else
                    NewAction gstrKEY_PRIVATEFILE, """" & strDestDir & strDestName & """"
                    CheckOverwritePrivateFile strDestDir & strDestName
                End If
            End If
            
            '
            'Si la información del archivo leída de SETUP.LST es el .EXE de la aplicación
            '(es decir, es el valor de la clave AppExe de la sección [Setup], se guarda
            'su ruta de acceso completa para su uso posterior
            '
            If strDestName = gstrAppExe Then
                '
                'Se usa para crear un icono del Administrador de programas en Form_Load de
                'SETUP1.FRM y para registrar la ruta de acceso por aplicación
                '
                gsDest.strAppDir = strDestDir
            End If

            'Caso especial para CTL3D32.DLL
            'nunca instalamos estos archivos bajo Windows 95 o NT4; sólo bajo Windows NT3.51
            If strDestName = mstrFILE_CTL3D32 Then
                If TreatAsWin95() Then
                    'No estamos ejecutando NT 3.51; no se instala este archivo.
                    intRC = IDIGNORE
                    LogNote ResolveResString(resCOMMON_CTL3D32NOTCOPIED, "|1", strDestName)
                    AbortAction
                End If
            End If
            
            'Caso especial para RICHED32.DLL
            'sólo instalamos este archivo bajo Windows 95, no bajo Windows NT (3.51 ó 4.0)
            If strDestName = mstrFILE_RICHED32 Then
                If Not IsWindows95() Then
                    'No estamos ejecutando Win95; no se instala este archivo.
                    intRC = IDIGNORE
                    LogNote ResolveResString(resCOMMON_RICHED32NOTCOPIED, "|1", strDestName)
                    AbortAction
                End If
            End If
            '
            ' Caso especial para AXDIST.EXE
            ' Si estamos en Win95 o NT4 y AXDIST.EXE está en la lista de instalación,
                ' necesitamos ejecutarlo cuando se complete setup1. AXDIST.EXE es un
                ' ejecutable autoextraíble que instala archivos especiales necesarios
            ' para la funcionalidad de Internet.
            '
            If UCase(strDestName) = gstrFILE_AXDIST Then
                '
                ' No se hace nada aquí si no estamos usando Win95 o NT4.
                '
                If Not TreatAsWin95() Then
                    'No estamos ejecutando Win95 o NT4; no se instala este archivo.
                    intRC = IDIGNORE
                    LogNote ResolveResString(resCOMMON_AXDISTNOTCOPIED, "|1", strDestName)
                    AbortAction
                    gfAXDist = False
                End If
            End If
            '
                ' Caso especial para AXDIST.EXE
            ' Si estamos en NT3.51 y WINt351.EXE está en la lista de instalación,
                ' necesitamos ejecutarlo cuando se complete setup1. WINt351.EXE es un
                ' ejecutable autoextraíble que instala archivos especiales necesarios
            ' para la funcionalidad de Internet.
            '
            If UCase(strDestName) = gstrFILE_WINT351 Then
                '
                ' No se hace nada aquí si no estamos usando NT3.51.
                '
                If TreatAsWin95() Then
                    'No estamos ejecutando NT3.51; no se instala este archivo.
                    intRC = IDIGNORE
                    LogNote ResolveResString(resCOMMON_WINT351NOTCOPIED, "|1", strDestName)
                    AbortAction
                    gfWINt351 = False
                End If
            End If
            
            strRegister = sFile.strRegister

            lThisFileSize = CalcFinalSize(sFile.lFileSize, sFile.strDestDir)

            '
            'El código siguiente intenta ahorrar tiempo al comprobar previamente
                'si se debe instalar o no un archivo antes de concatenar un archivo
                'dividido o antes de que VerInstallFile piense que implica una lectura
                'completa del archivo (para un archivo comprimido) como mínimo.
                'Básicamente, si ambos archivos tienen número de versión, se
                'compararán. Si un archivo tiene un número de versión y el otro no,
                'el que tiene número de versión se denominará "más reciente". Si
                'ningún archivo tiene número de versión, comparamos la fecha.
            '
            'Siempre se intenta obtener el número de versión del archivo de origen.
                'Si el archivo de información de la instalación no contiene ningún
                'número de versión (sSrcVerInfo.nMSHi = gintNOVERINFO), intentamos
                'leer el número de versión del archivo de origen. La lectura del
                'número de versión de un archivo dividido siempre fallará. Por eso
                'es conveniente incluir el número de versión para un archivo
                '(especialmente para los divididos) en el archivo de información
                'de la instalación (SETUP.LST)
            '
            fSrcVer = True
            sSrcVerInfo = sFile.sVerInfo
            If sSrcVerInfo.nMSHi = gintNOVERINFO Then
                fSrcVer = GetFileVerStruct(strSrcDir & strSrcName, sSrcVerInfo)
            End If

            '
            'Si existe un archivo de destino con información de la versión, se
                'compara su número de versión con el del archivo de origen.
            '
            If intRC <> IDIGNORE Then
                fRemoteReg = (sFile.strRegister = mstrREMOTEREGISTER)
                If GetFileVerStruct(strDestDir & strDestName, sDestVerInfo, fRemoteReg) = True Then
                    If fSrcVer = True Then
                        If IsNewerVer(sSrcVerInfo, sDestVerInfo) = False Then
                            '
                            'El archivo existente es más reciente que el que queremos instalar;
                            'se usará en su lugar el archivo existente
                            '
                            intRC = IDIGNORE
                            fFileWasUpToDate = True
                            DecideIncrementRefCount strDestDir & strDestName, sFile.fShared, sFile.fSystem, True
                            AddActionNote ResolveResString(resLOG_FILEUPTODATE)
                            CommitAction
                        End If
                    End If
                Else
                    '
                    'Si el archivo de destino no tiene información de la versión, copiaremos
                          'el archivo de origen si *tiene* alguna versión. Si ninguno de los
                          'archivos tiene número de versión, compararemos la fecha.
                    '
                    If fSrcVer = False Then
                        If sFile.varDate <= CVDate(FileDateTime(strDestDir & strDestName)) Then
                            If Err = 0 Then
                                '
                                'Aunque ni el archivo de origen ni el de destino contienen
                                          'información de la versión, el archivo existente tiene una
                                          'fecha posterior, por lo que la usaremos.
                                '
                                intRC = IDIGNORE
                                fFileWasUpToDate = True
                                DecideIncrementRefCount strDestDir & strDestName, sFile.fShared, sFile.fSystem, True
                                AddActionNote ResolveResString(resLOG_FILEUPTODATE)
                                CommitAction
                            Else
                                Err = 0
                            End If
                        End If
                    End If
                End If
            End If
            
            '
            'Si hemos decidido probar la copia y esto es el primer fragmento de un
                'archivo dividido, abriremos el archivo temporal usado para la
                'concatenación
            '
            If intRC <> IDIGNORE And fSplit = True Then
                mintConcatFile = OpenConcatFile()
                If mintConcatFile = -1 Then
                    'Falló la apertura y el usuario eligió ignorar el error
                    mintConcatFile = 0
                    intRC = IDIGNORE
                    AbortAction
                End If
            End If
        End If

        '
        'Si esto es un fragmento de un archivo dividido y vamos a probar la copia,
        'anexamos este fragmento del archivo de origen al final del archivo de
          'concatenación
        '
        If fSplit = True Then
            If intRC <> IDIGNORE Then
                intRC = ConcatSplitFile(strSrcDir & strSrcName)
                If intRC = IDIGNORE Then
                    AbortAction
                End If
            End If

            If intRC = IDIGNORE And mintConcatFile > 0 Then
                Close mintConcatFile
                mintConcatFile = 0
            End If

            fSplit = sFile.fSplit
        End If

        '
        'Si el archivo no estaba dividido o si es el último fragmento de un archivo dividido
        '
        If fSplit = False Then
            If mintConcatFile > 0 Then
                '
                'Si es el último fragmento de un archivo dividido, cerramos el archivo
                    'concatenado. En este momento, el archivo concatenado es una
                'representación real del archivo de origen deseado, por lo que
                'apuntamos a él en lugar de al fragmento del archivo dividido en el
                'medio de instalación
                '
                Close mintConcatFile
                strSrcDir = mstrConcatDrive
                strSrcName = mstrCONCATFILE
            End If

            '
            'Después de todo esto, si seguimos estando preparados para la copia, no
                'perdemos nada por intentarlo.
            '
            If intRC <> IDIGNORE Then
                ' CopyFile incrementará automáticamente el número de referencia y
                ' confirmará o anulará la acción actual.
                intRC = IIf(CopyFile(strSrcDir, strDestDir, strSrcName, strDestName, sFile.fShared, sFile.fSystem), 0, IDIGNORE)
            End If

            '
            'Guarda las rutas de acceso de ciertos archivos para su uso posterior,
            'si se instalaron correctamente o ya estaban en el sistema
            '
            If (intRC = 0 Or fFileWasUpToDate) Then
                Select Case strDestName
                    Case mstrFILE_AUTMGR32
                        '
                        'Se usa para crear un icono, si está instalado
                        '
                        gsDest.strAUTMGR32 = strDestDir & mstrFILE_AUTMGR32
                    Case mstrFILE_RACMGR32
                        '
                        'Se usa para crear un icono, si está instalado
                        '
                        gsDest.strRACMGR32 = strDestDir & mstrFILE_RACMGR32
                    'End Case
                End Select
            
                '
                'Si copiamos correctamente el archivo y se especificó información
                    'de registro en el archivo de información de la instalación, se
                    'guarda esa información en una matriz de forma que podamos
                    'registrar todos los archivos que lo requieran una vez copiados
                'todos los archivos.
                '
                If strRegister <> gstrNULL Then
                    Err = 0
                    ReDim Preserve msRegInfo(UBound(msRegInfo) + 1)
    
                    If Err > 0 Then
                        ReDim msRegInfo(0)
                    End If
    
                    msRegInfo(UBound(msRegInfo)).strFilename = strDestDir & strDestName
    
                    Select Case strRegister
                        Case mstrDLLSELFREGISTER, mstrEXESELFREGISTER, mstrTLBREGISTER, mstrVBLREGISTER
                            'No hay que hacer nada en particular
                        Case mstrREMOTEREGISTER
                            'Necesitamos buscar y analizar la correspondiente línea "RemoteX=..."
                            If Not ReadSetupRemoteLine(strSection, intIdx, msRegInfo(UBound(msRegInfo))) = True Then
                                MsgError ResolveResString(resREMOTELINENOTFOUND, "|1", strDestName, "|2", gstrINI_REMOTE & Format$(intIdx)), vbExclamation Or vbOKOnly, gstrTitle
                                ExitSetup frmSetup1, gintRET_FATAL
                            End If
                        Case Else
                            '
                            'Si la información de registro especificaba el nombre de un
                                    'archivo con información de registro (lo que suponemos si no
                                    'se especificó ninguna macro de registro), también suponemos
                                    'que, si no hay disponible información sobre la ruta de
                                    'acceso, este archivo de información de registro estará en
                                    'el mismo directorio que el archivo que registra
                            '
                            strRegister = ResolveDestDirs(strRegister)
                            If InStr(strRegister, gstrSEP_DIR) = 0 Then
                                strRegister = strDestDir & strRegister
                            End If
                        'End Case
                    End Select
    
                    msRegInfo(UBound(msRegInfo)).strRegister = strRegister
                End If
            
            End If

            '
            'Si creamos un archivo de concatenación temporal, lo quitamos
            '
            If mintConcatFile > 0 Then
                Kill mstrConcatDrive & mstrCONCATFILE
                mintConcatFile = 0
            End If
        End If

        strLastFile = sFile.strDestName

CSContinue:
        '
        'Si el archivo no estaba dividido, o si era el último fragmento de un archivo
          'dividido, actualizamos la barra de estado de la copia. Necesitamos hacer la
          'actualización, independientemente de que se haya copiado realmente un
          'archivo o no.
        '
        If sFile.fSplit = False Then
            glTotalCopied = glTotalCopied + lThisFileSize
            UpdateStatus frmCopy.picStatus, glTotalCopied / mlTotalToCopy
        End If

        '
        'Damos una oportunidad para que se procese el botón 'Cancelar' si se hizo clic en él
        '
        DoEvents
        intIdx = intIdx + 1
    Loop

    Err = 0
End Sub

'-----------------------------------------------------------
' Procedimiento: CreateOSProgramGroup
'
' Llama a CreateProgManGroup en Windows NT o llama a
' fCreateShellGroup en Windows 95
'-----------------------------------------------------------
'
Function fCreateOSProgramGroup(frm As Form, ByVal strFolderName As String, ByVal fRetOnErr As Boolean, Optional ByVal fLog As Boolean = True) As Boolean
    If TreatAsWin95() Then
        fCreateOSProgramGroup = fCreateShellGroup(strFolderName, fRetOnErr, fLog)
    Else
        CreateProgManGroup frm, strFolderName, fRetOnErr, fLog
        fCreateOSProgramGroup = True
    End If
End Function

'-----------------------------------------------------------
' Procedimiento: CreateOSLink
'
' Llama a CreateProgManItem en Windows NT o llama a
' CreateFolderLink en Windows 95.
'
' Si falta fLog, el valor predeterminado es True.
'-----------------------------------------------------------
'
Sub CreateOSLink(frm As Form, ByVal strGroupName As String, ByVal strLinkPath As String, ByVal strLinkArguments As String, ByVal strLinkName As String, Optional ByVal fLog)
    If IsMissing(fLog) Then
        fLog = True
    End If
    
    If TreatAsWin95() Then
        CreateShellLink strLinkPath, strGroupName, strLinkArguments, strLinkName, fLog
    Else
        '
        ' DDE no funcionará correctamente si intenta enviar a NT el nombre largo de
          ' archivo. Si está entre comillas, se pasarán por alto los parámetros. Si
          ' no hay ningún parámetro, puede utilizarse el nombre largo de archivo y
          ' se podría omitir la línea siguiente.
        '
        strLinkPath = GetShortPathName(strUnQuoteString(strLinkPath))
        CreateProgManItem frm, strGroupName, strLinkPath & " " & strLinkArguments, strLinkName, fLog
    End If
End Sub

'-----------------------------------------------------------
' Procedimiento: CreateProgManGroup
'
' Crea un nuevo grupo en el Administrador de programas de
' Windows si el nombre de grupo especificado no existe ya
'
' ENTRADA: [frm] - formulario que contiene una etiqueta 'lblDDE'
'     [strGroupName] - nombre de texto del grupo
'     [fRetOnErr]    - se pasa por alto
'     [fLog] - si se va a escribir o no en el archivo de registro
'                (si falta, el valor predeterminado es True)
'-----------------------------------------------------------
'
Sub CreateProgManGroup(frm As Form, ByVal strGroupName As String, ByVal fRetOnErr As Boolean, Optional ByVal fLog)
    '
    'Llama a una función DDE genérica de ProgMan con un indicador para agregar un grupo
    '
    If IsMissing(fLog) Then
        fLog = True
    End If
    
    'Realiza DDE para crear el grupo
    PerformDDE frm, strGroupName, gstrNULL, gstrNULL, mintDDE_GRPADD, fLog
End Sub

'-----------------------------------------------------------
' Procedimiento: CreateProgManItem
'
' Crea (o reemplaza) un icono del Administrador de programas
' en el grupo activo del Administrador de programas
'
' ENTRADA: [frm] - formulario que contiene una etiqueta 'lblDDE'
'     [strGroupName] - título del grupo donde aparecerá el icono.
'     [strCmdLine] - línea de comandos para el elemento o icono,
'                    por ejemplo: 'c:\miapl\miapl.exe'
'                    Nota: si esta ruta de acceso contiene espacios
'                      en blanco o comas, debe escribirse entre
'                      comillas para que Windows la interprete
'                      correctamente (vea AddQuotesToFN)
'     [strIconTitle] - título de texto para el icono
'     [fLog] - si se va a escribir o no en el archivo de registro
'                (el valor predeterminado es True si falta)
'
' CONDICIÓN PREVIA: ya se ha llamado a CreateProgManGroup. El nuevo
'               icono se creará en el último grupo creado.
'-----------------------------------------------------------
'
Sub CreateProgManItem(frm As Form, ByVal strGroupName As String, ByVal strCmdLine As String, ByVal strIconTitle As String, Optional ByVal fLog)
    '
    'Llamada a función DDE genérica de ProgMan con indicador para agregar
    'un elemento
    '
    If IsMissing(fLog) Then
        fLog = True
    End If
    PerformDDE frm, strGroupName, strCmdLine, strIconTitle, mintDDE_ITEMADD, fLog
End Sub

'-----------------------------------------------------------
' Procedimiento: fCreateShellGroup
'
' Crea un nuevo grupo de programas fuera de Inicio>Programas
' en el intérprete de comandos de Windows 95 si la carpeta
' especificada no existe ya.
'
' ENTRADA: [strFolderName] - nombre de texto de la carpeta.
'                      Este parámetro no puede contener
'                      barras inversas.
'                      ej.: "Mi aplicación" crea la carpeta
'                        Inicio>Programas>Mi aplicación
'     [fRetOnerr] - si esta rutina debe volver o no si hay un error
'                   al crear el grupo. Si es False, la instalación
'                   se anula y no vuelve. Establézcalo a True si el
'                   usuario puede hacer algo para corregir el error.
'                   Por ejemplo, se escribió un nombre de grupo en el
'                   cuadro de diálogo Elegir grupo de programas en
'                   lugar de llamar a esta rutina al crear el grupo
'                         Automatización remota en el cual el usuario no
'                         tiene ningún control.
'     [fLog] - si se va a escribir o no en el archivo de registro
'              (el valor predeterminado es True si falta)
'-----------------------------------------------------------
'
Function fCreateShellGroup(ByVal strFolderName As String, fRetOnErr As Boolean, Optional ByVal fLog) As Boolean
    If IsMissing(fLog) Then
        fLog = True
    End If

    ReplaceDoubleQuotes strFolderName
    
    If strFolderName = "" Then
        Exit Function
    End If

    If fLog Then
        NewAction gstrKEY_SHELLFOLDER, """" & strFolderName & """"
    End If

Retry:
    
    Dim fSuccess As Boolean
    fSuccess = OSfCreateShellGroup(strFolderName)
    If fSuccess Then
        If fLog Then
            CommitAction
        End If
    Else
        If gfNoUserInput Or (MsgError(ResolveResString(resCANTCREATEPROGRAMGROUP, "|1", strFolderName), vbRetryCancel Or vbExclamation, gstrTitle)) = IDCANCEL Then
            ExitSetup frmSetup1, gintRET_EXIT
            GoTo Retry
        End If
        '
        ' Determina si debemos volver de forma que el usuario pueda
        ' corregir la situación.
        '
        If Not fRetOnErr Then
            '
            ' Se vuelve para poder salir de la instalación.
            '
            GoTo Retry
        End If
    End If

    
    fCreateShellGroup = fSuccess
End Function

'-----------------------------------------------------------
' Procedimiento: CreateShellLink
'
' Crea (o reemplaza) un vínculo en Inicio>Programas o en
' cualquiera de sus subcarpetas inmediatas en el intérprete
' de comandos de Windows 95.
'
' ENTRADA: [strLinkPath] - ruta de acceso completa al destino del vínculo
'                     Ej.: 'c:\Archivos de programa\Mi aplicación\MiApl.exe"
'     [strLinkArguments] - argumentos de la línea de comandos para el vínculo
'                     Ej.: '-f -c "c:\Archivos de programa\Mi aplicación\MiApl.dat" -q'
'     [strLinkName] - título de texto para el vínculo
'     [fLog] - si se va a escribir o no en el archivo de registro
'              (el valor predeterminado es True si falta)
'
' SALIDA:
'   El vínculo se creará en la carpeta strGroupName

'-----------------------------------------------------------
'
Sub CreateShellLink(ByVal strLinkPath As String, ByVal strGroupName As String, ByVal strLinkArguments As String, ByVal strLinkName As String, Optional ByVal fLog)
    If IsMissing(fLog) Then
        fLog = True
    End If
    
    If fLog Then
        NewAction gstrKEY_SHELLLINK, """" & strUnQuoteString(strGroupName) & """" & ", " & """" & strUnQuoteString(strLinkName) & """"
    End If
    
    'ReplaceDoubleQuotes strLinkName
    strLinkName = strUnQuoteString(strLinkName)
    strLinkPath = strUnQuoteString(strLinkPath)
    

Retry:

    Dim fSuccess As Boolean
    fSuccess = OSfCreateShellLink(strGroupName & "", strLinkName, strLinkPath, strLinkArguments & "") 'la ruta no debe estar con comillas
    If fSuccess Then
        If fLog Then
            CommitAction
        End If
    Else
        Dim intMsgRet As Integer
        intMsgRet = MsgError(ResolveResString(resCANTCREATEPROGRAMICON, "|1", strLinkName), vbAbortRetryIgnore Or vbExclamation, gstrTitle)
        If gfNoUserInput Then
            intMsgRet = vbAbort
        End If
        Select Case intMsgRet
            Case vbAbort
                ExitSetup frmSetup1, gintRET_ABORT
                GoTo Retry
            Case vbRetry
                GoTo Retry
            Case vbIgnore
                If fLog Then
                    AbortAction
                End If
            'End Case
        End Select
    End If
End Sub

'-----------------------------------------------------------
' Función: DecideIncrementRefCount
'
' Incrementa el número de referencias de un archivo bajo 32 bits
' si se trata de un archivo compartido.
'
' ENTRADA: [strFullPath] - ruta completa del archivo para el número de referencias.
'                     Ejemplo:
'                     'C:\MIAPL\MIAPL.DAT'
'     [fShared] - si el archivo es compartido o privado
'     [fSystem] - es un archivo de sistema
'     [fFileAlreadyExisted] - si el archivo ya existía o no
'                             en el disco duro antes que
'                             nuestro programa de instalación
'-----------------------------------------------------------
'
Sub DecideIncrementRefCount(ByVal strFullPath As String, ByVal fShared As Boolean, ByVal fSystem As Boolean, ByVal fFileAlreadyExisted As Boolean)
    'El recuento de referencias tiene lugar tanto en Windows 95 como en Windows NT
    If fShared Or fSystem Then
        IncrementRefCount strFullPath, fFileAlreadyExisted
    End If
End Sub
            
'-----------------------------------------------------------
' Función: DetectFile
'
' Detecta si el archivo especificado existe. Si no se encuentra,
' se ofrece al usuario la oportunidad de anular, reintentar o
' ignorar la búsqueda del archivo. Esta llamada se utiliza, por
' ejemplo, para asegurarse de que en la unidad hay un disco con
' el nombre de archivo especificado antes de continuar.
'
' ENTRADA: [strFileName] - nombre del archivo que se va a detectar;
'                     normalmente debe incluir la ruta completa.
'                     Ejemplo: 'A:\MIAPL.DAT'
'
' Retorno: TRUE si se detectó el archivo, IDIGNORE si el
'          usuario eligió ignorar cuando no se pudo encontrar
'            el archivo o llama a ExitSetup al elegir anular
'-----------------------------------------------------------
'
Function DetectFile(ByVal strFilename As String) As Integer
    Dim strMsg As String

    DetectFile = True
                      
    Do While FileExists(strFilename) = False


        strMsg = ResolveResString(resCANTOPEN) & LS$ & strFilename
        Select Case MsgError(strMsg, MB_ABORTRETRYIGNORE Or MB_ICONEXCLAMATION Or IIf(gfNoUserInput, vbDefaultButton1, MB_DEFBUTTON2), gstrSETMSG)
            Case IDABORT
                ExitSetup frmCopy, gintRET_ABORT
            Case IDIGNORE
                DetectFile = IDIGNORE
                Exit Do
            'End Case
        End Select
    Loop
End Function


'-----------------------------------------------------------
' Procedimiento: EtchedLine
'
' Dibuja una línea 'grabada' con la forma especificada
' empezando en la ubicación X,Y pasada y de la longitud
' indicada. Las coordenadas están en el valor actual de
' ScaleMode del formulario pasado.
'
' ENTRADA: [frmEtch] - formulario en el que dibujar la línea
'     [intX1] - inicio horizontal de la línea
'     [intY1] - inicio vertical de la línea
'     [intLength] - longitud de la línea
'-----------------------------------------------------------
'
Sub EtchedLine(frmEtch As Form, ByVal intX1 As Integer, ByVal intY1 As Integer, ByVal intLength As Integer)
    Const lWHITE& = vb3DHighlight
    Const lGRAY& = vb3DShadow

    frmEtch.Line (intX1, intY1)-(intX1 + intLength, intY1), lGRAY
    frmEtch.Line (frmEtch.CurrentX + 5, intY1 + 20)-(intX1 - 5, intY1 + 20), lWHITE
End Sub

'-----------------------------------------------------------
' Procedimiento: ExeSelfRegister
'
' Ejecuta de forma síncrona el archivo pasado (que debe ser
' un archivo ejecutable que acepte el modificador /REGSERVER;
' por ejemplo, un .EXE de un componente ActiveX generado
' con VB5).
'
' ENTRADA: [strFileName] - archivo .EXE que se va a registrar
'-----------------------------------------------------------
'
Sub ExeSelfRegister(ByVal strFilename As String)
    Const strREGSWITCH$ = " /REGSERVER"

    Dim fShell As Integer

    '
    'Sale de forma síncrona y ejecuta el archivo .EXE con el
    'modificador de auto-registro
    '
    fShell = FSyncShell(AddQuotesToFN(strFilename) & strREGSWITCH, 7)
    frmSetup1.Refresh
End Sub

'-----------------------------------------------------------
' Procedimiento: ExitSetup
'
' Trata el cierre de la aplicación de instalación. Dependiendo
' del valor del parámetro intExitCode, puede preguntar al
' usuario y salir del subprocedimiento si el usuario elige
' cancelar el proceso de salida.
'
' ENTRADA: [frm] - formulario activo para descargar al salir
'     [intExitCode] - código que especifica la acción de salida
'-----------------------------------------------------------
'
Sub ExitSetup(frm As Form, intExitCode As Integer)
    Dim strMsg As String
    Dim strSilent As String

    On Error Resume Next
    '
    ' Si no estamos en modo desatendido o sms, se ofrece al
    ' usuario la posibilidad de reintentarlo, si es aplicable.
    '
    If Not gfNoUserInput Then
        Select Case intExitCode
            Case gintRET_EXIT
                '
                'Si el usuario eligió el botón Salir o Cancelar
                '
                If MsgWarning(ResolveResString(resASKEXIT), MB_ICONQUESTION Or MB_YESNO Or MB_DEFBUTTON2, gstrTitle) = IDNO Then
                    Exit Sub
                End If
            Case gintRET_ABORT
                '
                'Si el usuario eligió anular antes de una acción pendiente
                '
                strMsg = ResolveResString(resINCOMPLETE) & LS$ & ResolveResString(resQUITNOW) & LS$
                strMsg = strMsg & ResolveResString(resQUITSETUP)
                If MsgWarning(strMsg, MB_ICONQUESTION Or MB_YESNO Or IIf(gfNoUserInput, vbDefaultButton1, MB_DEFBUTTON2), gstrSETMSG) = IDNO Then
                    Exit Sub
                End If
            'End Case
        End Select
    End If

    'Anula cualquier acción pendiente
    While fWithinAction()
        AbortAction
    Wend
    
    Close

    '
    'Limpia los archivos temporales de VerInstallFile o la concatenación de archivos divididos
    '
    Kill mstrVerTmpName
    If mintConcatFile > 0 Then
        Close mintConcatFile
        Kill mstrConcatDrive & mstrCONCATFILE
    End If

    If frm.hWnd <> frmSetup1.hWnd Then
        Unload frm
    End If
    
    frmSetup1.SetFocus

    '
    'Ofrece el mensaje apropiado de fin, dependiendo del código de salida
    '
    Select Case intExitCode
        Case gintRET_EXIT, gintRET_ABORT
            gfSMSStatus = False
            strMsg = ResolveResString(resINTERRUPTED, "|1", gstrAppName) & LS$ & ResolveResString(resCANRUN, "|1", gstrAppName)
            MsgWarning strMsg, MB_OK Or MB_ICONSTOP, gstrTitle
        Case gintRET_FATAL
            gfSMSStatus = False
            MsgError ResolveResString(resERROR, "|1", gstrAppName), MB_OK Or MB_ICONSTOP, gstrTitle
        Case gintRET_FINISHEDSUCCESS
            gfSMSStatus = True
            '
            ' No registra este mensaje en SMS, ya que sólo es una confirmación.
            '
            gfDontLogSMS = True
            MsgFunc ResolveResString(resSUCCESS, "|1", gstrAppName), MB_OK, gstrTitle
            
            If IsWindowsNT4WithoutSP2() Then
                'Se recomienda que el usuario se actualize a NT 4.0 SP2
                gfDontLogSMS = True
                MsgWarning ResolveResString(resNT4WithoutSP2), MB_OK Or MB_ICONINFORMATION, gstrTitle
            End If
        Case Else
            strMsg = ResolveResString(resINTERRUPTED, "|1", gstrAppName) & LS$ & ResolveResString(resCANRUN, "|1", gstrAppName)
            MsgWarning strMsg, MB_OK Or MB_ICONSTOP, gstrTitle
        'End Case
    End Select

    'Detiene el registro
    DisableLogging
    
    ' Limpia una instalación anulada
    If (intExitCode = gintRET_FINISHEDSUCCESS) Then
        'La instalación terminó correctamente; ya se deben haber limpiado
        'los archivos temporales. No hay nada más que hacer.
    Else
        'Se anuló la instalación por alguna razón
        If (gstrAppRemovalEXE <> "") Then
            Dim nErrorLevel As Integer
            Select Case intExitCode
                Case gintRET_FATAL
                    nErrorLevel = APPREMERR_FATAL
                Case gintRET_EXIT
                    nErrorLevel = APPREMERR_USERCANCEL
                Case gintRET_ABORT
                    nErrorLevel = APPREMERR_NONFATAL
                Case Else
                    nErrorLevel = APPREMERR_FATAL
                'End Case
            End Select
        
            '
            ' No queremos registrar este mensaje en sms porque sólo es un
            ' mensaje de confirmación.
            '
            gfDontLogSMS = True
            MsgFunc ResolveResString(resLOG_ABOUTTOREMOVEAPP), vbInformation Or vbOKOnly, gstrTitle
            
            Err = 0
            '
            ' Preparado para ejecutar el instalador. Determina si es una
            ' instalación desatendida o no.
            '
            If gfSilent Then
                strSilent = gstrSilentLog
            Else
                strSilent = gstrNULL
            End If
            
            Shell GetAppRemovalCmdLine(gstrAppRemovalEXE, gstrAppRemovalLog, strSilent, gfSMS, nErrorLevel, True), vbNormalFocus
            If Err Then
                MsgError Error$ & LS$ & ResolveResString(resLOG_CANTRUNAPPREMOVER), MB_ICONEXCLAMATION Or MB_OK, gstrTitle
            End If

            'Como el programa para quitar aplicaciones intentará eliminar este
                'programa y todos nuestros archivos de tiempo de ejecución, debemos
                'salir lo antes posible (de lo contrario, el programa para quitar
                'aplicaciones no podrá quitar estos archivos)
        End If
        
        'Nota: no eliminamos el archivo de registro si se produce un error.
        'El EXE para quitar aplicaciones lo hará si es necesario.
        
    End If
    
    Unload frmSetup1

    If gfSMS = True Then
        WriteMIF gstrMIFFile, gfSMSStatus, gstrSMSDescription
    End If

    'Termina el programa
    End
End Sub

'-----------------------------------------------------------
' Función: ProcessCommandLine
'
' Procesa los argumentos de la línea de comandos
'
' SALIDA: llena los parámetros pasados por referencia
' según sea apropiado
'-----------------------------------------------------------
'
Sub ProcessCommandLine(ByVal strCommand As String, ByRef fSilent As Boolean, ByRef strSilentLog As String, ByRef fSMS As Boolean, ByRef strMIFFile As String, ByRef strSrcPath As String, ByRef strAppRemovalLog As String, ByRef strAppRemovalEXE As String)
    Dim fErr As Boolean
    Dim intAnchor As Integer
    
    Const strTemp$ = ""
    
    strSrcPath = ""
    strAppRemovalLog = ""
    
    strCommand = Trim$(strCommand)
    
    '
    ' En primer lugar, comprueba si ésta es una instalación desatendida
    ' (/s o -s en la línea de comandos, seguido de un nombre de archivo
    ' de registro) y establece las variables en consecuencia.
    '
    ' Si está diseñando una instalación desatendida, debe agregarse al
    ' comando setup.exe el parámetro /s o -s de la línea de comandos.
    ' Se pasará automáticamente a setup1 como primer parámetro.
    '
    ' El nombre de archivo que sigue al parámetro /s o -s debe incluir
    ' el nombre completo de la ruta de acceso.
    '
    intAnchor = InStr(LCase(strCommand), gstrSwitchPrefix1 & gstrSILENTSWITCH)
    If intAnchor = 0 Then
        intAnchor = InStr(LCase(strCommand), gstrSwitchPrefix2 & gstrSILENTSWITCH)
    End If
    If intAnchor > 0 Then
        fSilent = True
        strCommand = Trim(Mid(strCommand, intAnchor + 2))
        strSilentLog = strExtractFilenameArg(strCommand, fErr)
        If fErr Then GoTo BadCommandLine
    Else
        fSilent = False
    End If
    '
    ' Después comprueba si ésta es una instalación desatendida SMS.
    ' Si la instalación se inició con el modificador /q o -q, es una
    ' instalación desatendida SMS. /q o -q debe ir seguido del nombre
    ' del archivo MIF de SMS en el que escribir información de estado.
    ' Al llamar a setup.exe se pasa el nombre del exe de la aplicación
    ' como el nombre de su archivo MIF (por ejemplo, /q MiProg.Exe).
    ' Setup.exe tomará este nombre de archivo y lo convertirá en
    ' c:\windows\MiProg.MIF (suponiendo que windows es el directorio
    ' "c:\windows") y lo pasará a setup1. Esencialmente, setup.exe
    ' tomará cualquier nombre de archivo que se le pase a continuación
    ' del modificador q, quitará la ruta de acceso y la extensión y,
    ' después, le agregará la ruta del directorio windows y la extensión
    ' MIF para crear el nombre del archivo MIF. No importa si pasa o no
    ' la ruta completa de acceso a setup.exe.
    '
    intAnchor = InStr(LCase(strCommand), gstrSwitchPrefix1 & gstrSMSSWITCH)
    If intAnchor = 0 Then
        intAnchor = InStr(LCase(strCommand), gstrSwitchPrefix2 & gstrSMSSWITCH)
    End If
    If intAnchor > 0 Then
        fSMS = True
        strCommand = Trim(Mid(strCommand, intAnchor + 2))
        strMIFFile = strExtractFilenameArg(strCommand, fErr)
        If fErr Then GoTo BadCommandLine
    Else
        fSMS = False
    End If
    
    '
    ' Esperamos encontrar el directorio de origen, el nombre y
    ' la ruta de acceso del archivo de registro, y el nombre y
    ' la ruta de acceso del ejecutable para quitar aplicaciones,
    ' separados sólo por espacios en blanco
    '
    strSrcPath = strExtractFilenameArg(strCommand, fErr)
    If fErr Then GoTo BadCommandLine
    
    strAppRemovalLog = strExtractFilenameArg(strCommand, fErr)
    If fErr Then GoTo BadCommandLine
    

    strAppRemovalEXE = strExtractFilenameArg(strCommand, fErr)
    If fErr Then GoTo BadCommandLine
        
    ' El archivo de registro y el ejecutable para quitar aplicaciones
    ' deben existir
    If Not FileExists(strAppRemovalLog) Then
        GoTo BadAppRemovalLog
    End If
    
    If Not FileExists(strAppRemovalEXE) Then
        GoTo BadAppRemovalEXE
    End If
    
    ' Última comprobación: no debe haber nada más en la línea de comandos
    strCommand = Trim$(strCommand)
    If strCommand <> "" Then
        GoTo BadCommandLine
    End If
    
    Exit Sub
    
BadAppRemovalLog:
    MsgError ResolveResString(resCANTFINDAPPREMOVALLOG, "|1", strAppRemovalLog), MB_ICONEXCLAMATION Or MB_OK, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
    
BadAppRemovalEXE:
    MsgError ResolveResString(resCANTFINDAPPREMOVALEXE, "|1", strAppRemovalEXE), MB_ICONEXCLAMATION Or MB_OK, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
    
BadCommandLine:
    MsgError ResolveResString(resBADCOMMANDLINE), MB_ICONEXCLAMATION Or MB_OK, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Sub

'-----------------------------------------------------------
' Función: GetDrivesAllocUnit
'
' Obtiene la unidad de asignación del tamaño mínimo de
' archivo para la unidad especificada
'
' ENTRADA: [strDrive] - unidad para la cual obtener la unidad
'          de asignación
'
' Retorno: unidad mínima de asignación de la unidad o -1 si
'          no se pudo determinar este valor
'-----------------------------------------------------------
'
Function GetDrivesAllocUnit(ByVal strDrive As String) As Long
    Dim strCurDrive As String
    Dim lAllocUnit As Long

    On Error Resume Next

    '
    'Guarda la unidad actual
    '
    strCurDrive = Left$(CurDir$, 2)

    '
    'anexa un signo de dos puntos al final de la especificación
    'de unidad si no se suministra ninguna
    '
    If InStr(strDrive, gstrCOLON) = 0 Or Len(strDrive) > 2 Then
        strDrive = Left$(strDrive, 1) & gstrCOLON
    End If

    '
    'Cambia a la unidad para la cual determinar la unidad de asignación.
    'La API AllocUnit() devuelve este valor para la unidad actual
    'únicamente
    '
    ChDrive strDrive

    '
    'Si hubo un error al tener acceso a la unidad especificada, se devuelve
    'un error de indicador.
    'La API It AllocUnit() también puede devolver -1 si se produce otro fallo
    '
    If Err <> 0 Or (strDrive <> Left$(CurDir$, 2)) Then
        lAllocUnit = -1
    Else
        lAllocUnit = AllocUnit()
        If Err <> 0 Then
            lAllocUnit = -1
        End If
    End If

    If lAllocUnit = -1 Then
        MsgError Error$ & LS$ & ResolveResString(resALLOCUNIT) & strDrive, MB_ICONEXCLAMATION, gstrTitle
        If gfSMS Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If

    GetDrivesAllocUnit = lAllocUnit

    '
    'Se restaura a la unidad original
    '
    ChDrive strCurDrive

    Err = 0
End Function

'-----------------------------------------------------------
' Función: GetFileName
'
' Devuelve la parte de nombre de archivo de una ruta de acceso
'
'-----------------------------------------------------------
'
Function GetFileName(ByVal strPath As String) As String
    Dim strFilename As String
    Dim iSep As Integer
    
    strFilename = strPath
    Do
        iSep = InStr(strFilename, gstrSEP_DIR)
        If iSep = 0 Then iSep = InStr(strFilename, gstrCOLON)
        If iSep = 0 Then
            GetFileName = strFilename
            Exit Function
        Else
            strFilename = Right(strFilename, Len(strFilename) - iSep)
        End If
    Loop
End Function

'-----------------------------------------------------------
' Función: GetFileSize
'
' Determina el tamaño (en bytes) del archivo especificado
'
' ENTRADA: [strFileName] - nombre del archivo del cual se
'          va a obtener el tamaño
'
' Retorno: tamaño del archivo en bytes, o -1 si se produce
'          un error
'-----------------------------------------------------------
'
Function GetFileSize(strFilename As String) As Long
    On Error Resume Next

    GetFileSize = FileLen(strFilename)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function

'-----------------------------------------------------------
' Función: GetAppRemovalCmdLine
'
' Devuelve los argumentos correctos de la línea de comandos
' (incluyendo la ruta de acceso al ejecutable que se usará
' al llamar al ejecutable para quitar aplicaciones)
'
' ENTRADA: [strAppRemovalEXE] - ruta de acceso y nombre completos
'                        del EXE para quitar aplicaciones
'     [strAppRemovalLog] - ruta de acceso y nombre completos
'                        del archivo de registro para quitar
'                        aplicaciones
'     [strSilentLog] - ruta de acceso y nombre completos del
'                        archivo donde se registrarán los mensajes
'                        en modo desatendido. Si es una cadena
'                        vacía, se desactivará el modo desatendido
'                        para la desinstalación.
'     [fSMS] - valor Boolean. Si es True, hemos estado realizando
'              una instalación SMS y debemos indicar al desinstalador
'              que también haga una desinstalación SMS. SMS es
'              Microsoft Systems Management Server.
'     [nErrorLevel] - nivel de error:
'                        APPREMERR_NONE - ningún error
'                        APPREMERR_FATAL - error fatal
'                        APPREMERR_NONFATAL - error no fatal; el usuario eligió anular
'                        APPREMERR_USERCANCEL - el usuario eligió cancelar (ningún error)
'     [fWaitForParent] - es True si el programa para quitar aplicaciones debe esperar
'                        a que termine el primario (este proceso) antes de empezar a
'                        quitar archivos. De lo contrario, quizá no pueda quitar el
'                        archivo ejecutable de este proceso, dependiendo del tiempo.
'                        El valor predeterminado es False, si no se especifica nada.
'-----------------------------------------------------------
'
Function GetAppRemovalCmdLine(ByVal strAppRemovalEXE As String, ByVal strAppRemovalLog, ByVal strSilentLog As String, ByVal fSMS As Boolean, ByVal nErrorLevel As Integer, Optional fWaitForParent)
    Dim strEXE As String
    Dim strLog As String
    Dim strSilent As String
    Dim strErrLevel As String
    Dim strForce As String
    Dim strWait As String
    Dim strSMS As String

    If IsMissing(fWaitForParent) Then
        fWaitForParent = False
    End If
    
    strEXE = AddQuotesToFN(strAppRemovalEXE)
    strLog = "-n " & """" & GetLongPathName(strAppRemovalLog) & """"
    If gfSilent And strSilentLog <> gstrNULL Then
        strSilent = "/s " & """" & strSilentLog & """"
    Else
        strSilent = gstrNULL
    End If
    
    strSMS = IIf(fSMS, " /q ", gstrNULL)
    
    strErrLevel = IIf(nErrorLevel <> APPREMERR_NONE, "-e " & Format(nErrorLevel), "")
    If nErrorLevel <> APPREMERR_NONE Then
        strForce = " -f"
    End If
    If fWaitForParent Then
        Dim curProcessId As Currency
        Dim Wrap As Currency
        Dim lProcessId As Long
        Dim cProcessId As Currency
        
        Wrap = 2 * (CCur(&H7FFFFFFF) + 1)

        'Siempre lo imprime como un valor Long sin signo
        lProcessId = GetCurrentProcessId()
        cProcessId = lProcessId
        If cProcessId < 0 Then cProcessId = cProcessId + Wrap

        strWait = " -w " & str(cProcessId)
    End If
    
    GetAppRemovalCmdLine = strEXE & " " & strLog & " " & strSilent & " " & strSMS & strErrLevel & strForce & strWait
End Function

'-----------------------------------------------------------
' Función: IncrementRefCount
'
' Incrementa el número de referencias de un archivo en el
' registro, de forma que se pueda quitar correctamente si
' el usuario elige quitar esta aplicación.
'
' ENTRADA: [strFullPath] - ruta de acceso y nombre completos
'                          del archivo
'     [fFileAlreadyExisted] - indica si el archivo indicado
'                             ya existía en el disco duro
'-----------------------------------------------------------
'
Sub IncrementRefCount(ByVal strFullPath As String, ByVal fFileAlreadyExisted As Boolean)
    Dim strSharedDLLsKey As String
    strSharedDLLsKey = RegPathWinCurrentVersion() & "\SharedDLLs"
    
    'Siempre debemos usar el nombre largo para el archivo, de forma que podamos
    'identificar de forma única y exacta el archivo en el registro.
    strFullPath = GetLongPathName(strFullPath)
    
    'Obtiene el número de referencia actual para este archivo
    Dim fSuccess As Boolean
    Dim hKey As Long
    fSuccess = RegCreateKey(HKEY_LOCAL_MACHINE, strSharedDLLsKey, "", hKey)
    If fSuccess Then
        Dim lCurRefCount As Long
        If Not RegQueryRefCount(hKey, strFullPath, lCurRefCount) Then
            'No hay ningún número de referencia actual para este archivo
            If fFileAlreadyExisted Then
                'Si no había ningún número de referencia, pero el archivo se
                'encontró en el disco duro, significa una de dos cosas:
                '  1) Este archivo se incluye con el sistema operativo
                '  2) Este archivo lo instaló un programa de instalación
                '     antiguo que no hacía recuento de referencias
                'En cualquier caso, lo más conservador es suponer que alguna
                'aplicación necseita el archivo, lo que significa que debe
                'tener un número de referencia de 1 como mínimo. De esta
                'forma, nuestro programa para quitar aplicaciones no
                'eliminará este archivo.
                lCurRefCount = 1

            Else
                lCurRefCount = 0
            End If
        End If
        
        'Incrementa el número en el registro
        fSuccess = RegSetNumericValue(hKey, strFullPath, lCurRefCount + 1, False)
        If Not fSuccess Then
            GoTo DoErr
        End If
        RegCloseKey hKey
    Else
        GoTo DoErr
    End If
    
    Exit Sub
    
DoErr:
    'Ya se debe haber mostrado un mensaje de error al usuario
    Exit Sub
End Sub

'-----------------------------------------------------------
' Función: InitDiskInfo
'
' Se le llama antes de calcular el espacio en disco para
' inicializar valores utilizados o determinados al calcular
' el espacio en disco necesario.
'-----------------------------------------------------------
'
Sub InitDiskInfo()
    '
    'Inicializa la "tabla" de unidades usadas y la matriz de espacio en disco
    '
    gstrDrivesUsed = gstrNULL
    Erase gsDiskSpace

    mlTotalToCopy = 0

    '
    'Obtiene la unidad y el directorio para archivos temporales
    '
    mstrConcatDrive = UCase$(Environ$(gstrTMP_DIR))
    If mstrConcatDrive = gstrNULL Then
        mstrConcatDrive = UCase$(Environ$(gstrTEMP_DIR))
    End If
    AddDirSep mstrConcatDrive

    If mstrConcatDrive <> gstrNULL Then
        If CheckDrive(mstrConcatDrive, ResolveResString(resTEMPDRIVE)) = False Then
            mstrConcatDrive = gstrNULL
        Else
            '
            'Si encontramos una unidad temporal y ésta está "preparada", se
            'agrega a la tabla de unidades utilizadas
            '
            gstrDrivesUsed = Left$(mstrConcatDrive, 1)
            ReDim Preserve gsDiskSpace(1)
            gsDiskSpace(1).lAvail = GetDiskSpaceFree(mstrConcatDrive)
            gsDiskSpace(1).lMinAlloc = GetDrivesAllocUnit(mstrConcatDrive)
        End If
    End If
End Sub

'-----------------------------------------------------------
' Función: IsDisplayNameUnique
'
' Determina si un nombre de presentación especificado para
' registrar el ejecutable para quitar aplicaciones es único
' o no. Este nombre es el título que se muestra al usuario
' en la aplicación Agregar o quitar programas del Panel de
' control de Windows 95.
'
' ENTRADA: [hkeyAppRemoval] - clave libre para la ruta de
'                       acceso del registro que contiene
'                       entradas para quitar aplicaciones
'     [strDisplayName] - nombre cuya unicidad se va a comprobar
'
' Retorno: True si el nombre de presentación ya está en uso;
'          de lo contrario, False
'-----------------------------------------------------------
'
Function IsDisplayNameUnique(ByVal hkeyAppRemoval As Long, ByVal strDisplayName As String) As Boolean
    Dim lIdx As Long
    Dim strSubkey As String
    Dim strDisplayNameExisting As String
    Const strKEY_DISPLAYNAME$ = "DisplayName"
    
    IsDisplayNameUnique = True
    
    lIdx = 0
    Do
        Select Case RegEnumKey(hkeyAppRemoval, lIdx, strSubkey)
            Case ERROR_NO_MORE_ITEMS
                'No hay más claves; debe ser única
                Exit Do
            Case ERROR_SUCCESS
                'Tenemos una clave para algún programa para quitar aplicaciones.
                'Se compara su nombre de presentación con el nuestro
                Dim hkeyExisting As Long
                
                If RegOpenKey(hkeyAppRemoval, strSubkey, hkeyExisting) Then
                    If RegQueryStringValue(hkeyExisting, strKEY_DISPLAYNAME, strDisplayNameExisting) Then
                        If strDisplayNameExisting = strDisplayName Then
                            'Hay una coincidencia con un nombre de presentación existente
                            IsDisplayNameUnique = False
                            RegCloseKey hkeyExisting
                            Exit Do
                        End If
                    End If
                    RegCloseKey hkeyExisting
                End If
            Case Else
                'Error; debemos suponer que se único. Puede que ocurra un error
                'más tarde al intentar agregarlo al registro
                Exit Do
            'End Case
        End Select
        lIdx = lIdx + 1
    Loop
End Function

'-----------------------------------------------------------
' Función: IsNewerVer
'
' Compara dos estructuras de versiones de archivos y determina
' si la versión del archivo de origen es más reciente (mayor)
' que la del archivo de destino. Esto se utiliza para
' determinar si hay que instalar o no un archivo
'
' ENTRADA: [sSrcVer] - información de la versión del archivo de origen
'     [sDestVer] - información de la versión del archivo de destino
'
' Retorno: True si el archivo de origen es más reciente que el de destino;
'          de lo contrario, False
'-----------------------------------------------------------
'
Function IsNewerVer(sSrcVer As VERINFO, sDestVer As VERINFO) As Integer
    IsNewerVer = False

    If sSrcVer.nMSHi > sDestVer.nMSHi Then GoTo INVNewer
    If sSrcVer.nMSHi < sDestVer.nMSHi Then GoTo INVOlder
    
    If sSrcVer.nMSLo > sDestVer.nMSLo Then GoTo INVNewer
    If sSrcVer.nMSLo < sDestVer.nMSLo Then GoTo INVOlder
    
    If sSrcVer.nLSHi > sDestVer.nLSHi Then GoTo INVNewer
    If sSrcVer.nLSHi < sDestVer.nLSHi Then GoTo INVOlder
    
    If sSrcVer.nLSLo > sDestVer.nLSLo Then GoTo INVNewer

    GoTo INVOlder

INVNewer:
    IsNewerVer = True
INVOlder:
End Function

'-----------------------------------------------------------
' Función: IsValidDestDir
'
' Determina si el directorio de destino especificado en la
' clave "DefaultDir" de la sección [Setup] de SETUP.LST o
' un directorio de destino especificado por el usuario
' no es un subdirectorio del directorio de origen.
'
' Notas: [gstrSrcPath] - señala al directorio de origen
'        [strDestDir] - señala al directorio de destino
'
' Retorno: True si el directorio de destino es una ubicación
'          válida; de lo contrario, False
'-----------------------------------------------------------
'
Function IsValidDestDir(strDestDir As String) As Integer
    Dim strMsg As String

    '
    ' Ambas rutas de acceso, strDestDir y gstrSrcPath, tienen *siempre*
    ' el formato 'X:\' o 'X:\NOMBRE_DIR\'.
    '
    If InStr(strDestDir, gstrSrcPath) > 0 Then
        IsValidDestDir = False
        strMsg = ResolveResString(resDIRSPECIFIED) & LF$ & strDestDir & LF$ & ResolveResString(resSAMEASSRC)
        MsgFunc strMsg, MB_OK Or MB_ICONEXCLAMATION, gstrTitle
    Else
        IsValidDestDir = True
    End If
End Function

'-----------------------------------------------------------
' Función: MakePath
'
' Crea la ruta de acceso del directorio especificado
'
' ENTRADA: [strDirName] - nombre de la ruta de acceso del directorio
'                         que se va a crear
'     [fAllowIgnore] - si se permitirá o no al usuario pasar
'                      por alto los errores encontrados. Si es
'                      False, la función  sólo vuelve si tiene
'                      éxito. Si falta, de manera predeterminada
'                      es True.
'
' Retorno: True si tiene éxito; False si hay un error y el
'          usuario eligió ignorar. (La función no vuelve si
'          el usuario elige ANULAR/CANCELAR cuando hay un error.)
'-----------------------------------------------------------
'
Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore) As Boolean
    If IsMissing(fAllowIgnore) Then
        fAllowIgnore = True
    End If
    
    Do
        If MakePathAux(strDir) Then
            MakePath = True
            Exit Function
        Else
            Dim strMsg As String
            Dim iRet As Integer
            
            strMsg = ResolveResString(resMAKEDIR) & LF$ & strDir
            iRet = MsgError(strMsg, IIf(fAllowIgnore, MB_ABORTRETRYIGNORE, MB_RETRYCANCEL) Or MB_ICONEXCLAMATION Or MB_DEFBUTTON2, gstrSETMSG)
            '
            ' si estamos en modo desatendido, no podemos
            ' continuar. El MsgError anterior se encargó
            ' de escribir la entrada de registro desatendida.
            '
            If gfNoUserInput = True Then
                ExitSetup frmCopy, gintRET_FATAL
            End If
            
            Select Case iRet
                Case IDABORT, IDCANCEL
                    ExitSetup frmCopy, gintRET_ABORT
                Case IDIGNORE
                    MakePath = False
                    Exit Function
                Case IDRETRY
                'End Case
            End Select
        End If
    Loop
End Function

'----------------------------------------------------------
' Procedimiento: MoveAppRemovalFiles
'
' Mueve el archivo de registro para quitar aplicaciones al
' directorio de la aplicación y registra el ejecutable en
' el sistema operativo.
'----------------------------------------------------------
Sub MoveAppRemovalFiles(ByVal strGroupName As String)
    Dim strNewAppRemovalLogName As String
    
    'Busca en el directorio de la aplicación un nombre único para el
    'archivo de registro para quitar aplicaciones
    
    '...Primero prueba con la extensión predeterminada
    strNewAppRemovalLogName = gstrDestDir & mstrFILE_APPREMOVALLOGBASE & mstrFILE_APPREMOVALLOGEXT
    If FileExists(strNewAppRemovalLogName) Then
        '...Después prueba a incrementar las extensiones integrales
        Dim iExt As Integer
        Do
            If iExt > 999 Then
                GoTo CopyErr
            End If
            

            strNewAppRemovalLogName = gstrDestDir & mstrFILE_APPREMOVALLOGBASE & gstrSEP_EXT & Format(iExt, "000")
            If Not FileExists(strNewAppRemovalLogName) Then
                Exit Do 'Se encontró un nombre único
            Else
                iExt = iExt + 1
            End If
        Loop
    End If
    
    
    
    On Error GoTo CopyErr
    FileCopy gstrAppRemovalLog, strNewAppRemovalLogName
    
    'Ahora necesitamos empezar a registrar en el nuevo archivo de
    'registro, de forma que quede registrada la creación del icono
    'para quitar aplicaciones bajo NT.
    EnableLogging strNewAppRemovalLogName
    
    On Error GoTo 0
    If Not RegisterAppRemovalEXE(gstrAppRemovalEXE, strNewAppRemovalLogName, strGroupName) Then
        If TreatAsWin95() Then
            MsgError ResolveResString(resCANTREGISTERAPPREMOVER), MB_ICONEXCLAMATION Or MB_OK, gstrTitle
        Else
            MsgError ResolveResString(resCANTCREATEAPPREMOVALICON), MB_ICONEXCLAMATION Or MB_OK, gstrTitle
        End If
        ExitSetup frmSetup1, gintRET_FATAL
    End If
    
    'Ahora podemos eliminar el archivo de registro original, ya que no
    'tenemos una referencia al mismo, y empezar a utilizar el nuevo
    'archivo de registro
    On Error Resume Next
    Kill gstrAppRemovalLog
    
    'Este archivo de registro temporal para quitar aplicaciones no se
    'debe utilizar
    gstrAppRemovalLog = strNewAppRemovalLogName
    gfAppRemovalFilesMoved = True
    
    Exit Sub
    
CleanUpOnErr:
    On Error Resume Next
    Kill strNewAppRemovalLogName
    On Error GoTo 0
    MsgError ResolveResString(resCANTCOPYLOG, "|1", gstrAppRemovalLog), vbExclamation Or vbOKOnly, gstrTitle
    ExitSetup Screen.ActiveForm, gintRET_FATAL
    
CopyErr:
    Resume CleanUpOnErr
End Sub

'-----------------------------------------------------------
' Función: OpenConcatFile
'
' Abre un archivo para que sea el destino de la concatenación
' de dos o más archivos de origen que (normalmente) se han
' dividido en varios discos.
'
' Retorno: el controlador del archivo que se usará para la
'          concatenación si la apertura se realizó
'          correctamente o -1 si la apertura falló y el
'          usuario eligió ignorar el error.
'-----------------------------------------------------------
'
Function OpenConcatFile() As Integer
    Dim intFileNum As Integer
    Dim strMsg As String

    On Error Resume Next

    Do
        Kill mstrConcatDrive & mstrCONCATFILE
        Err = 0

        intFileNum = FreeFile
        Open mstrConcatDrive & mstrCONCATFILE For Binary Access Write As intFileNum

        If Err > 0 Then
            strMsg = ResolveResString(resNOCREATE) & LS$ & mstrConcatDrive & mstrCONCATFILE
            strMsg = strMsg & LS$ & ResolveResString(resNOTPROTECT)
            Select Case MsgError(strMsg, MB_ABORTRETRYIGNORE Or MB_ICONEXCLAMATION Or IIf(gfNoUserInput, vbDefaultButton1, MB_DEFBUTTON2), gstrSETMSG)
                Case IDABORT
                    ExitSetup frmCopy, gintRET_ABORT
                Case IDIGNORE
                    OpenConcatFile = -1
                    Exit Function
                'End Case
            End Select
        End If
    Loop While Err > 0

    OpenConcatFile = intFileNum
End Function

'-----------------------------------------------------------
' Procedimiento: ParseDateTime
'
' Igual que CDate con un argumento de cadena, excepto que
' pasa por alto los valores actuales de localización. Esto
' es importante porque SETUP.LST siempre utiliza el mismo
' formato para las fechas.
'
' ENTRADA: [strDate] - cadena que representa la fecha en el
'                 formato mm/dd/aa o mm/dd/aaaa
' SALIDA: la fecha que representa strDate
'-----------------------------------------------------------
'
Function ParseDateTime(ByVal strDateTime As String) As Date
    Const strDATESEP$ = "/"
    Const strTIMESEP$ = ":"
    Const strDATETIMESEP$ = " "
    Dim iMonth As Integer
    Dim iDay As Integer
    Dim iYear As Integer
    Dim iHour As Integer
    Dim iMinute As Integer
    Dim iSecond As Integer
    Dim iPos As Integer
    Dim vTime As Date
    
    iPos = InStr(strDateTime, strDATESEP)
    If iPos = 0 Then GoTo Err
    iMonth = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iPos = InStr(strDateTime, strDATESEP)
    If iPos = 0 Then GoTo Err
    iDay = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iPos = InStr(strDateTime, strDATETIMESEP)
    If iPos = 0 Then GoTo SkipTime
    iYear = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    vTime = TimeSerial(0, 0, 0)
    
    iPos = InStr(strDateTime, strTIMESEP)
    If iPos = 0 Then GoTo SkipTime
    iHour = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iPos = InStr(strDateTime, strTIMESEP)
    If iPos = 0 Then GoTo SkipTime
    iMinute = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iSecond = Val(strDateTime)
    
    vTime = TimeSerial(iHour, iMinute, iSecond)
    
SkipTime:
    
    If iYear < 100 Then iYear = iYear + 1900
    
    ParseDateTime = DateSerial(iYear, iMonth, iDay) + vTime
    
    Exit Function
    
Err:
    Error 13 'Error de tipo, igual que un CDate intríseco provoca un error
End Function

'-----------------------------------------------------------
' Procedimiento: PerformDDE
'
' Realiza una operación DDE del Administrador de programas
' según especifique el indicador intDDE y los parámetros
' pasados.
' Las posibles operaciones son las siguientes:
'
'   mintDDE_ITEMADD:  Agregar un icono al grupo activo
'   mintDDE_GRPADD:   Crear un grupo en el Administrador de programas
'
' ENTRADA: [frm] - formulario que contiene una etiqueta llamada 'lblDDE'
'     [strGroup] - nombre del grupo para crear o insertar el icono
'     [strTitle] - título del icono o del grupo
'     [strCmd] - línea de comandos para el icono o el elemento
'                que se va a agregar
'     [intDDE] - acción DDE de ProgMan que se va a realizar
'-----------------------------------------------------------
'
Sub PerformDDE(frm As Form, ByVal strGroup As String, ByVal strCmd As String, ByVal strTitle As String, ByVal intDDE As Integer, ByVal fLog As Boolean)
    Const strCOMMA$ = ","
    Const strRESTORE$ = ", 1)]"
    Const strACTIVATE$ = ", 5)]"
    Const strENDCMD$ = ")]"
    Const strSHOWGRP$ = "[ShowGroup("
    Const strADDGRP$ = "[CreateGroup("
    Const strREPLITEM$ = "[ReplaceItem("
    Const strADDITEM$ = "[AddItem("

    Dim intIdx As Integer        'variable de bucle

    SetMousePtr gintMOUSE_HOURGLASS

    '
    'Inicializa la conversación DDE con el Administrador de programas
    'de Windows en modo manual (.LinkMode = 2), donde el control de
    'destino no se actualiza automáticamente. Establece el tiempo de
    'espera DDE durante 10 segundos. El bucle que hay alrededor de
    'DoEvents() sirve para dejar tiempo para que se procese DDEExecute.
    '

    Dim intRetry As Integer
    For intRetry = 1 To 20
        On Error Resume Next
        frm.lblDDE.LinkTopic = "PROGMAN|PROGMAN"
        If Err = 0 Then
            Exit For
        End If
        DoEvents
    Next intRetry
        
    frm.lblDDE.LinkMode = 2
    For intIdx = 1 To 10
      DoEvents
    Next
    frm.lblDDE.LinkTimeout = 100

    On Error Resume Next

    If Err = 0 Then
        Select Case intDDE
            Case mintDDE_ITEMADD
                '
                ' El elemento se creará en el grupo llamado strGroup
                '
                ' Escribe la acción en el archivo de registro
                '
                If fLog Then
                    NewAction gstrKEY_PROGMANITEM, """" & strUnQuoteString(strGroup) & """" & ", " & """" & strUnQuoteString(strTitle) & """"
                End If
                '
                ' Fuerza que el grupo strGroup sea el grupo activo. Additem sólo
                ' pone iconos en el grupo activo.
                '
                #If 0 Then
                    frm.lblDDE.LinkExecute strSHOWGRP & strGroup & strACTIVATE
                #Else
                    ' BUG #5-30466,stephwe,10/96: strShowGRP no parece funcionar si ProgMan está minimizado.
                    '  : strADDGRP funciona bien y no importa si ya existe.
                    frm.lblDDE.LinkExecute strADDGRP & strGroup & strENDCMD
                #End If
                frm.lblDDE.LinkExecute strREPLITEM & strTitle & strENDCMD
                Err = 0
                frm.lblDDE.LinkExecute strADDITEM & strCmd & strCOMMA & strTitle & String$(3, strCOMMA) & strENDCMD
            Case mintDDE_GRPADD
                '
                ' Escribe la acción en el archivo de registro
                '
                If fLog Then
                    NewAction gstrKEY_PROGMANGROUP, """" & strUnQuoteString(strGroup) & """"
                End If
                frm.lblDDE.LinkExecute strADDGRP & strGroup & strENDCMD
                frm.lblDDE.LinkExecute strSHOWGRP & strGroup & strRESTORE
            'End Case
        End Select
    End If

    
    '
    'Desconecta el vínculo DDE
    '

    frm.lblDDE.LinkMode = 0
    frm.lblDDE.LinkTopic = ""


    SetMousePtr gintMOUSE_DEFAULT

    If fLog Then
        CommitAction
    End If
    
    
    Err = 0
End Sub

'-----------------------------------------------------------
' Procedimiento: PromptForNextDisk
'
' Si el medio de origen es extraíble o es una conexión de
' red, se pide al usuario que inserte el número de disco
' especificado que contiene el nombre de archivo que se
' usará para determinar que se ha insertado el disco
' correcto.
'
' ENTRADA: [intDiskNum] - número de disco que se va a insertar
'     [strDetectFile] - archivo que se va a buscar para asegurarse
'                       de que se insertó el disco correcto
'
' Notas: [gstrSrcPath] - se usa para identificar la unidad de origen
'-----------------------------------------------------------
'
Sub PromptForNextDisk(ByVal intDiskNum As Integer, ByVal strDetectFile As String)
    Static intDrvType As Integer

    Dim intRC As Integer
    Dim strMsg As String
    Dim strDrive As String
    Dim strMultDirBaseName As String
    Dim strDetectPath As String

    On Error Resume Next

    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    '
    'Obtiene la unidad de origen y, si no lo hemos determinado todavía,
    'obtiene el tipo de unidad de origen
    '
    
    strDrive = Left$(gstrSrcPath, 2)
    If intDrvType = 0 Then
        If IsUNCName(strDrive) Then
            intDrvType = intDRIVE_REMOTE
            strDrive = gstrSrcPath
        Else
            intDrvType = GetDriveType(Asc(strDrive) - 65)
        End If
    End If

    While SrcFileMissing(gstrSrcPath, strDetectFile, intDiskNum) = True
        Select Case intDrvType
            Case 0, intDRIVE_REMOVABLE, intDRIVE_CDROM
                strMsg = ResolveResString(resINSERT) & LF$ & ResolveResString(resDISK) & Format$(intDiskNum)
                strMsg = strMsg & ResolveResString(resINTO) & strDrive
            Case intDRIVE_REMOTE
                strMsg = ResolveResString(resCHKCONNECT) & strDrive
            Case intDRIVE_FIXED
                If DirExists(gstrSrcPath & strMultDirBaseName & Format(intDiskNum)) = True Then
                    strDetectPath = gstrSrcPath & strMultDirBaseName & Format(intDiskNum)
                Else
                    strDetectPath = gstrSrcPath
                End If
                strMsg = ResolveResString(resCOMMON_CANTFINDSRCFILE, "|1", strDetectPath & gstrSEP_DIR & strDetectFile)
            'End Case
        End Select

        Beep
        intRC = MsgFunc(strMsg, MB_OKCANCEL Or MB_ICONEXCLAMATION, gstrSETMSG)
        '
        ' Siempre falla si estamos en modo desatendido o sms.
        '
        If intRC = IDCANCEL Or gfNoUserInput Then
            ExitSetup frmCopy, gintRET_EXIT
        End If
    Wend

    gintCurrentDisk = intDiskNum
End Sub
Function SrcFileMissing(ByVal strSrcDir As String, ByVal strSrcFile As String, ByVal intDiskNum As Integer) As Boolean
'-----------------------------------------------------------
' Función: SrcFileMissing
'
' Intenta encontrar el archivo strSrcFile buscando primero
' en el directorio strSrcDir y después en el directorio
' DISK(x+1) si ya existe.
'
' ENTRADA: [strSrcDir] - directorio o ruta de acceso donde
'                        debe estar el archivo.
'     [strSrcFile] - archivo que estamos buscando.
'     [intDiskNum] - número de disco donde esperamos que
'                    esté el archivo.
'
' Retorno: True si no se encuentra el archivo; de lo contrario, False
'-----------------------------------------------------------
    Dim fFound As Boolean
    Dim strMultDirBaseName As String
    
    fFound = False
    
    AddDirSep strSrcDir
    '
    ' Primero comprueba si está en el directorio principal
    ' de origen. Esto ocurrirá si alguien copió el contenido
    ' de todos los disquetes a un único directorio del disco
    ' duro. Debemos permitir que esto funcione.
    '
    ' Esta prueba también nos permitirá saber si el usuario
    ' insertó el disquete erróneo o si una conexión de red
    ' no está disponible.
    '
    If FileExists(strSrcDir & strSrcFile) = True Then
        fFound = True
        GoTo doneSFM
    End If
    '
    ' Después prueba el subdirectorio el DISK(x) del directorio
    ' principal de origen. Esto ocurrirá si los disquetes
    ' se copiaron a directorios llamados DISK1, DISK2,
    ' DISK3,..., DISKN, etc.
    '
    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    If FileExists(strSrcDir & ".." & gstrSEP_DIR & strMultDirBaseName & Format(intDiskNum) & gstrSEP_DIR & strSrcFile) = True Then
        fFound = True
        GoTo doneSFM
    End If
    
doneSFM:
    SrcFileMissing = Not fFound
End Function
'-----------------------------------------------------------
' Función: ReadIniFile
'
' Lee un valor de la sección o clave especificada del
' archivo .INI indicado
'
' ENTRADA: [strIniFile] - nombre del archivo .INI que se
'                         va a leer
'     [strSection] - sección donde se encontró la clave
'     [strKey] - nombre de la clave de la que obtener el valor
'
' Retorno: valor terminado en algo distinto de cero de la clave
' del archivo .INI
'-----------------------------------------------------------
'
Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer

    '
    'Si se lee correctamente el archivo .INI, elimina cualquier cero
    'posterior devuelto por la API de Windows GetPrivateProfileString
    '
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strSection, strKey, gstrNULL, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = gstrNULL
    End If
End Function

'-----------------------------------------------------------
' Procedimiento: ReadSetupFileLine
'
' Lee la clave 'FileX=' solicitada de la sección especificada
' del archivo de información de la instalación (SETUP.LST).
'
' ENTRADA: [strSection] - nombre de la sección que se leerá de SETUP.LST,
'                    por ejemplo: "Files"
'     [intFileNum] - índice de número de archivo que se va a leer
'
' SALIDA: [sFile] - variable de tipo FILEINFO que, después de analizarse,
'                contiene información para el archivo descrito.
'
' Retorno: True si la información solicitada se leyó correctamente;
'          de lo contrario, False
'
' Notas: las líneas del archivo de información de la instalación
'        tienen el siguiente formato:
'
'        #,[SPLIT],SrcName,DestName,DestDir,Register,
'        Date,Size,Version
'
'        [#] - número de disco donde se encuentra este archivo
'        [SPLIT] - opcional, determina si es un fragmento de un
'                  archivo dividido. El último fragmento no
'                  especifica esta clave
'        [SrcName] - nombre de archivo del medio de instalación
'        [DestName] - nombre de archivo que se usará al copiar
'
'        (Para los archivos divididos, la siguiente información sólo
'        es necesaria para el *primer* fragmento)
'
'        [DestDir] - nombre de directorio o macro que especifica
'                    el directorio de destino
'        [Register] - nombre de archivo de información de registro o macro
'                     que especifica la acción de registro del archivo
'        [Date] - fecha del archivo de origen
'        [Size] - tamaño del archivo de origen
'        [Version] - opcional; cadena de número de versión
'        [Reserved] - debe estar vacío; de lo contrario, se producirá un error
'        [ProgIcon] - título del icono, si hay alguno
'        [ProgCmdLine] - línea de comandos para el icono, si hay alguno
'-----------------------------------------------------------
'
Function ReadSetupFileLine(ByVal strSection As String, ByVal intFileNum As Integer, sFile As FILEINFO) As Integer
    Static strSplitName As String
    Const CompareBinary = 0

    Dim strLine As String
    Dim strMsg As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim fDone As Integer
    Dim fErr As Boolean
    Dim strVersion As String
    Dim strFilename As String

    ReadSetupFileLine = False

    sFile.fSystem = False
    sFile.fShared = False
    
    '
    ' Lee la línea solicitada; si no puede leerla (strLine = gstrNULL) sale
    '
    strLine = ReadIniFile(gstrSetupInfoFile, strSection, gstrINI_FILE & Format$(intFileNum))
    If strLine = gstrNULL Then
        Exit Function
    End If

    '
    ' Obtiene el número de disco
    '
    intOffset = intGetNextFldOffset(1, strLine, gstrCOMMA, CompareBinary)
    sFile.intDiskNum = Val(Left$(strLine, intOffset - 1))
    If sFile.intDiskNum < 1 Then
        GoTo RSFLError
    End If

    '
    'Es un fragmento de un archivo dividido (no es el último fragmento)
    '
    intAnchor = intOffset + 1
    
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        sFile.fSplit = IIf(Mid$(strLine, intAnchor, intOffset - intAnchor) = gstrNULL, False, True)
    Else
        GoTo RSFLError
    End If

    '
    'nombre del archivo de origen; se asegura de que no es un nombre UNC
    '
    intAnchor = intOffset + 1
    sFile.strSrcName = strExtractFilenameItem(strLine, intAnchor, fErr)
    If fErr Then GoTo RSFLError
    If IsUNCName(sFile.strSrcName) = True Then GoTo RSFLError
    intAnchor = intAnchor + 1 'Salta la coma
    
    '
    'nombre del archivo de destino; se asegura de que no es un nombre UNC
    '
    sFile.strDestName = strExtractFilenameItem(strLine, intAnchor, fErr)
    If fErr Then GoTo RSFLError
    If IsUNCName(sFile.strDestName) = True Then GoTo RSFLError
    If Mid$(strLine, intAnchor, 1) = gstrCOMMA Then
        If IsUNCName(sFile.strDestName) = True Then
            GoTo RSFLError
        End If
        intAnchor = intAnchor + 1 'Salta la coma
    Else
        '
        'Si no hay ningún separador de listas después del nombre de archivo
          'de destino, debe ser un fragmento de un archivo dividido
        '
        If strSplitName = gstrNULL Then
            GoTo RSFLError
        Else
            sFile.strDestDir = gstrNULL
            fDone = True
        End If
    End If
    
    strFilename = GetFileName(sFile.strDestName)

    '
    'Se asegura de que los archivos DIVIDIDOS de SETUP.LST terminan correctamente;
    'para ello se comprueba que todos los nombres de archivos de destino después
    'de la primera línea SPLIT son idénticos, hasta e incluyendo el nombre del
    'archivo de destino de la siguiente línea que *no* es SPLIT.
    '
    If sFile.fSplit = True Then
        If strSplitName = gstrNULL Then
            strSplitName = sFile.strDestName
        Else
            If strSplitName <> sFile.strDestName Then
                GoTo RSFLError
            End If
        End If
    Else
        If strSplitName <> gstrNULL And strSplitName <> sFile.strDestName Then
            GoTo RSFLError
        Else
            strSplitName = gstrNULL

        End If
    End If

    If fDone = True Then
        GoTo RSFLDone
    End If

    '
    'analiza y resuelve el directorio de destino
    '
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        Dim strInitialDestDir As String
        strInitialDestDir = Mid$(strLine, intAnchor, intOffset - intAnchor)
        If InStr(strInitialDestDir, gstrWINSYSDESTSYSFILE) Then
            sFile.fSystem = True
        End If
        If InStr(strInitialDestDir, gstrDAODEST) Then
            '
            ' Caso especial para destinos de DAO. Si hay algún archivo
            ' de DAO, necesitamos agregar posteriormente información
            ' especial de registro de DAO. gfRegDAO nos pide que lo hagamos.
            '
            gfRegDAO = True
        End If
        sFile.strDestDir = ResolveDestDir(strInitialDestDir)
        If sFile.strDestDir <> "?" Then
            sFile.strDestDir = ResolveDir(sFile.strDestDir, False, False)
            If sFile.strDestDir = gstrNULL Or IsUNCName(sFile.strDestDir) Then
                GoTo RSFLError
            End If
        End If
    Else
        GoTo RSFLError
    End If

    '
    'información de registro de archivo
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        sFile.strRegister = Mid$(strLine, intAnchor, intOffset - intAnchor)
    Else
        GoTo RSFLError
    End If

    '
    'Extrae el tipo compartido de archivo
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    sFile.fShared = False
    If intOffset > 0 Then
        Dim strShareType As String
        strShareType = Mid$(strLine, intAnchor, intOffset - intAnchor)
        Select Case strShareType
            Case mstrPRIVATEFILE
                sFile.fShared = False
            Case mstrSHAREDFILE
                If sFile.fSystem Then
                    'Un archivo no puede ser a la vez de sistema y compartido
                    GoTo RSFLError
                End If
                
                sFile.fShared = True
            Case Else
                GoTo RSFLError
            'End Case
        End Select
    End If
    
    '
    'Extrae la fecha del archivo y la convierte a un Variant de fecha
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        If IsDate(Mid$(strLine, intAnchor, intOffset - intAnchor)) = True Then
            sFile.varDate = ParseDateTime(Mid$(strLine, intAnchor, intOffset - intAnchor))
        Else
            GoTo RSFLError
        End If
    End If

    '
    'Obtiene el tamaño del archivo
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        sFile.lFileSize = Val(Mid$(strLine, intAnchor, intOffset - intAnchor))
    Else
        GoTo RSFLError
    End If

    '
    ' Obtiene el número de versión; de lo contrario, indica que no hay
    ' información de la versión
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        strVersion = Trim(Mid$(strLine, intAnchor, intOffset - intAnchor))
        If strVersion = "" Then
            sFile.sVerInfo.nMSHi = gintNOVERINFO
        Else
            PackVerInfo strVersion, sFile.sVerInfo
        End If
    Else
        GoTo RSFLError
    End If
    
    '
    ' El siguiente campo está reservado para una versión futura.
    ' Si contiene información, se producirá un error. Utilizar
    ' este campo para cualquier otro fin puede hacer que falle
    ' la instalación en versiones futuras de Visual Basic.
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        If Len(strUnQuoteString(Mid(strLine, intAnchor, intOffset - intAnchor))) > 0 Then
            GoTo RSFLError
        End If
    Else
        GoTo RSFLError
    End If
    '
    ' Obtiene el título para el icono del programa, si hay alguno.
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    '
    ' Pasa por alto este campo para RacMgr.EXE y AutMgr32.EXE. Se tratan de
    ' forma diferente. Vea Form_Load()
    '
    If strFilename = mstrFILE_RACMGR32 Or strFilename = mstrFILE_AUTMGR32 Then
        sFile.strProgramIconTitle = ""
    Else
        If intOffset > 0 Then
            sFile.strProgramIconTitle = Trim(Mid(strLine, intAnchor, intOffset - intAnchor))
        Else
            GoTo RSFLError
        End If
    End If
    '
    ' Obtiene la línea de comandos para el icono del programa, si hay alguno.
    ' Observe que éste es el último elemento de la lista. No debe haber ninguna
    ' coma después de este elemento, pero lo comprobaremos por si acaso.
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    '
    ' Pasa por alto este campo para RacMgr.EXE y AutMgr32.EXE. Se tratan de manera
    ' diferente. Vea Form_Load()
    '
    If strFilename = mstrFILE_RACMGR32 Or strFilename = mstrFILE_AUTMGR32 Then
        sFile.strProgramIconCmdLine = ""
    Else
        If intOffset > 0 Then
            '
            ' Hay una coma al final de este campo. Se usa.
            '
            sFile.strProgramIconCmdLine = Trim(Mid(strLine, intAnchor, intOffset - intAnchor))
        Else
            sFile.strProgramIconCmdLine = Trim(Mid(strLine, intAnchor))
        End If
    End If
    
RSFLDone:
    ReadSetupFileLine = True
    Exit Function

RSFLError:
    strMsg = gstrSetupInfoFile & LS$ & ResolveResString(resINVLINE) & LS$
    strMsg = strMsg & ResolveResString(resSECTNAME) & strSection & LF$ & strLine
    MsgError strMsg, MB_ICONSTOP, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Function

'-----------------------------------------------------------
' Procedimiento: ReadSetupRemoteLine
'
' Lee la clave 'RemoteX=' solicitada de la sección
' especificada del archivo de información de la instalación
' (SETUP.LST).
'
' ENTRADA: [strSection] - nombre de la sección que se va a leer de SETUP.LST,
'                    por ejemplo: "Files"
'     [intFileNum] - índice de número remoto que se va a leer
'
' SALIDA: [rInfo] - variable de tipo REGINFO que, una vez analizada,
'                contiene la información para la línea
'                descrita.
'
' Retorno: True si la información solicitada se leyó correctamente;
'          de lo contrario, False
'
' Notas: las líneas del servidor remoto del archivo de información
'        de la instalación tienen el siguiente formato:
'
'        dirección,protocolo,nivel-autentificación
'
'        [dirección] - dirección de la red del servidor, si se conoce
'        [protocolo] - nombre de protocolo de red, si se conoce
'        [nivel autentificación] - nivel de autentificación (o 0 para
'                                  el valor predeterminado)
'-----------------------------------------------------------
'
Function ReadSetupRemoteLine(ByVal strSection As String, ByVal intFileNum As Integer, rInfo As REGINFO) As Integer
    Dim strLine As String
    Dim strMsg As String
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim fErr As Boolean

    ReadSetupRemoteLine = False

    '
    'Lee la línea solicitada; si no puede leerla (strLine = gstrNULL) sale
    '
    strLine = ReadIniFile(gstrSetupInfoFile, strSection, gstrINI_REMOTE & Format$(intFileNum))
    If strLine = gstrNULL Then
        Exit Function
    End If

    '
    'Obtiene la dirección de red
    '
    intAnchor = 1
    fErr = False
    If Mid$(strLine, intAnchor, 1) = gstrCOMMA Then
        rInfo.strNetworkAddress = ""
    Else
        rInfo.strNetworkAddress = strExtractFilenameItem(strLine, intAnchor, fErr)
    End If
    If fErr Then GoTo RSRLError
    intAnchor = intAnchor + 1 'Salta la coma

    '
    'Obtiene el protocolo de red
    '
    If Mid$(strLine, intAnchor, 1) = gstrCOMMA Then
        rInfo.strNetworkProtocol = ""
    Else
        rInfo.strNetworkProtocol = strExtractFilenameItem(strLine, intAnchor, fErr)
    End If
    If fErr Then GoTo RSRLError
    intAnchor = intAnchor + 1 'Salta la coma

    '
    'Obtiene el nivel de autentificación (debe ser un único dígito
    '  entre 0..6)
    '
    Const intMaxAuthentication = 6
    Dim strAuthentication As String
    
    strAuthentication = Mid$(strLine, intAnchor, 1)
    If Len(strAuthentication) <> 1 Then GoTo RSRLError
    If (Asc(strAuthentication) < Asc("0")) Or (Asc(strAuthentication) > Asc("9")) Then GoTo RSRLError
    rInfo.intAuthentication = Val(strAuthentication)
    If rInfo.intAuthentication > intMaxAuthentication Then GoTo RSRLError
    '
    ' ¿Es DCOM o Automatización remota?
    '
    intAnchor = InStr(intAnchor + 1, strLine, gstrCOMMA)
    If intAnchor > 0 Then
        rInfo.fDCOM = (Trim(Mid$(strLine, intAnchor + 1)) = gstrDCOM)
    End If
    
    ReadSetupRemoteLine = True
    Exit Function

RSRLError:
    strMsg = gstrSetupInfoFile & LS$ & ResolveResString(resINVLINE) & LS$
    strMsg = strMsg & ResolveResString(resSECTNAME) & strSection & LF$ & strLine
    MsgError strMsg, MB_ICONSTOP, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Function

'-----------------------------------------------------------
' Función: RegCloseKey
'
' Cierra una clave abierta del registro.
'
' Retorno: True si tiene éxito; de lo contrario, False.
'-----------------------------------------------------------
'
Function RegCloseKey(ByVal hKey As Long) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    RegCloseKey = (lResult = ERROR_SUCCESS)
End Function

'-----------------------------------------------------------
' Función: RegCreateKey
'
' Abre (crea si ya existe) una clave en el registro del sistema.
'
' ENTRADA: [hkey] - el primario de HKEY.
'     [lpszSubKeyPermanent] - la primera parte de la subclave de
'         'hkey' que se creará o abrirá. El programa para quitar
'         aplicaciones (sólo de 32 bits) nunca eliminará ninguna
'         parte de esta subclave. NO puede ser una cadena vacía ("").
'     [lpszSubKeyRemovable] - la subclave de hkey\lpszSubKeyPermanent
'         que se creará o abrirá. Si se quita la aplicación (sólo
'         en 32 bits), se eliminará todo este subárbol si está vacío
'         en el momento de quitar la aplicación. Si este parámetro
'         es una cadena vacía (""), la entrada no se registrará.
'
' SALIDA: [phkResult] - HKEY de la clave recién creada o abierta.
'
' Retorno: True si la clave se creó o abrió correctamente; de lo
'   contrario, False
'   Si tiene éxito, phkResult se establece al controlador de la clave.
'
'-----------------------------------------------------------
Function RegCreateKey(ByVal hKey As Long, ByVal lpszSubKeyPermanent As String, ByVal lpszSubKeyRemovable As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    Dim fLog As Boolean
    Dim strSubKeyFull As String

    On Error GoTo 0

    If lpszSubKeyPermanent = "" Then
        RegCreateKey = False 'Error: lpszSubKeyPermanent no deber ser = ""
        Exit Function
    End If
    
    If Left$(lpszSubKeyRemovable, 1) = "\" Then
        lpszSubKeyRemovable = Mid$(lpszSubKeyRemovable, 2)
    End If

    If lpszSubKeyRemovable = "" Then
        fLog = False
    Else
        fLog = True
    End If
    
    If lpszSubKeyRemovable <> "" Then
        strSubKeyFull = lpszSubKeyPermanent & "\" & lpszSubKeyRemovable
    Else
        strSubKeyFull = lpszSubKeyPermanent
    End If
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGKEY, _
          """" & strHkey & "\" & lpszSubKeyPermanent & """" _
            & ", " & """" & lpszSubKeyRemovable & """"
    End If

    lResult = OSRegCreateKey(hKey, strSubKeyFull, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegCreateKey = True
        If fLog Then
            CommitAction
        End If
        AddHkeyToCache phkResult, strHkey & "\" & strSubKeyFull
    Else
        RegCreateKey = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

'-----------------------------------------------------------
' Función: RegDeleteKey
'
' Elimina una clave existente del registro del sistema.
'
' Retorno: True si tiene éxito; de lo contrario, False
'-----------------------------------------------------------
'
Function RegDeleteKey(ByVal hKey As Long, ByVal lpszSubKey As String) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegDeleteKey(hKey, lpszSubKey)
    RegDeleteKey = (lResult = ERROR_SUCCESS)
End Function

'-----------------------------------------------------------
' Procedimiento: RegEdit
'
' Llama a REGEDIT para agregar al registro del sistema la
' información del archivo especificado. Si el archivo .REG
' requiere información de la ruta de acceso en base al
' directorio de destino indicado por el usuario, necesitará
' escribir y llamar a una rutina de corrección de .REG antes
' de realizar el registro siguiente.
'
' ADVERTENCIA: no se recomienda el uso de esta funcionalidad
' ADVERTENCIA: bajo Win32, ya que el programa para quitar
' ADVERTENCIA: aplicaciones no permite deshacer cambios que
' ADVERTENCIA: se realizan como resultado de llamar a
' ADVERTENCIA: REGEDIT en un archivo .REG arbitrario.
' ADVERTENCIA: En su lugar, se recomienda que utilice las
' ADVERTENCIA: funciones RegCreateKey(), RegOpenKey(),
' ADVERTENCIA: RegSetStringValue(), etc. de este módulo.
' ADVERTENCIA: Esto crea entradas en el archivo de registro
' ADVERTENCIA: para quitar aplicaciones, lo que permite
' ADVERTENCIA: deshacer dichos cambios.
'
' ENTRADA: [strRegFile] - nombre del archivo que contiene
'                         información de registro
'-----------------------------------------------------------
'
Sub RegEdit(ByVal strRegFile As String)
    Const strREGEDIT$ = "REGEDIT /S "

    Dim fShellOK As Integer

    On Error Resume Next

    If FileExists(strRegFile) = True Then
        'Como regedit es una aplicación de 16 bits, no acepta dobles
        'comillas alrededor del nombre de archivo. Así, strRegFile contiene
        'espacios en blanco, la única forma de hacer que funcione es pasar a
        'regedit la versión reducida de la ruta del nombre de archivo.
        strRegFile = GetShortPathName(strRegFile)
        
        fShellOK = FSyncShell(strREGEDIT & strRegFile, 7)
        frmSetup1.Refresh
    Else
        MsgError ResolveResString(resCANTFINDREGFILE, "|1", strRegFile), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup frmSetup1, gintRET_FATAL
    End If

    Err = 0
End Sub

' Función: RegEnumKey
'
' Enumera las subclaves de una clave abierta del registro
' (devuelve la subclave número "i" de hkey, si existe)
'
' Retorno:
'   ERROR_SUCCESS si tiene éxito. strSubkeyName se establece al nombre de la subclave.
'   ERROR_NO_MORE_ITEMS si no hay más subclaves (sólo en 32 bits)
'   cualquier otra cosa - error
'
Function RegEnumKey(ByVal hKey As Long, ByVal i As Long, strKeyName As String) As Long
    Dim strResult As String
    
    strResult = String(300, " ")
    RegEnumKey = OSRegEnumKey(hKey, i, strResult, Len(strResult))
    strKeyName = StripTerminator(strResult)
End Function
'-----------------------------------------------------------
' Procedimiento: RegisterDAO
'
' Es necesario agregar claves especiales al registro si está
' instalado DAO. Esta rutina agrega esas claves.
'
' Tenga en cuenta que estas claves no se desinstalarán.
'
Sub RegisterDAO()
    Const strDAOKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}"
    Const strDAOKeyVal = "OLE 2.0 Link"
    Const strDAOInprocHandlerKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}\InprocHandler"
    Const strDAOInprocHandlerKeyVal = "ole2.dll"
    Const strDAOProgIDKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}\ProgID"
    Const strDAOProgIDKeyVal = "Access.OLE2Link"
    
    Dim hKey As Long
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOKey, "", hKey) Then
        '
        ' RegCreateKey muestra un error si algo va mal.
        '
        GoTo REGDAOError
    End If
    '
    ' Establece el valor de la clave
    '
    If Not RegSetStringValue(hKey, "", strDAOKeyVal, False) Then
        '
        ' RegSetStringValue muestra un error si algo va mal.
        '
        GoTo REGDAOError
    End If
    '
    ' Cierra la clave
    '
    RegCloseKey hKey
    '
    ' Repite el mismo proceso para las otras dos claves.
    '
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOInprocHandlerKey, "", hKey) Then GoTo REGDAOError
    If Not RegSetStringValue(hKey, "", strDAOInprocHandlerKeyVal, False) Then GoTo REGDAOError
    RegCloseKey hKey
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOProgIDKey, "", hKey) Then GoTo REGDAOError
    If Not RegSetStringValue(hKey, "", strDAOProgIDKeyVal, False) Then GoTo REGDAOError
    RegCloseKey hKey

    Exit Sub
        
REGDAOError:
    '
    ' Ya se deben haber mostrado mensajes de error.
    '
    ExitSetup frmSetup1, gintRET_FATAL
        
End Sub
'-----------------------------------------------------------
' Procedimiento: RegisterFiles
'
' Recorre la lista (matriz) de archivos para registrar el
' que se creó en la función CopySection y registra cada
' archivo cuando es necesario
'
' Notas: msRegInfo() matriz creada por la función CopySection
'-----------------------------------------------------------
'
Sub RegisterFiles()
    Const strEXT_EXE$ = "EXE"

    Dim intIdx As Integer
    Dim intLastIdx As Integer
    Dim strFilename As String
    Dim strMsg As String

    On Error Resume Next

    '
    'Obtiene el número de elementos para registrar; si no hay
    'ninguno, podemos salir de aquí
    '
    intLastIdx = UBound(msRegInfo)
    If Err > 0 Then
        GoTo RFCleanup
    End If

    For intIdx = 0 To intLastIdx
        strFilename = msRegInfo(intIdx).strFilename

        Select Case msRegInfo(intIdx).strRegister
            Case mstrDLLSELFREGISTER
                Dim intDllSelfRegRet As Integer
                Dim intErrRes As Integer
                Const FAIL_OLE = 2
                Const FAIL_LOAD = 3
                Const FAIL_ENTRY = 4
                Const FAIL_REG = 5
                
                NewAction gstrKEY_DLLSELFREGISTER, """" & strFilename & """"
                
RetryDllSelfReg:
                Err = 0
                intErrRes = 0
                intDllSelfRegRet = DLLSelfRegister(strFilename)
                If Err Then
                    intErrRes = resCOMMON_CANTREGUNEXPECTED
                Else
                    Select Case intDllSelfRegRet
                        Case 0
                            'Bien; todo es correcto
                        Case FAIL_OLE
                            intErrRes = resCOMMON_CANTREGOLE
                        Case FAIL_LOAD
                            intErrRes = resCOMMON_CANTREGLOAD
                        Case FAIL_ENTRY
                            intErrRes = resCOMMON_CANTREGENTRY
                        Case FAIL_REG
                            intErrRes = resCOMMON_CANTREGREG
                        Case Else
                            intErrRes = resCOMMON_CANTREGUNEXPECTED
                        'End Case
                    End Select
                End If
                
                If intErrRes Then
                    'Hubo algún tipo de error
                    
                    'Registra la versión más técnica del mensaje de error;
                    'si se mostrara al usuario final, le confundiría
                    LogError ResolveResString(intErrRes, "|1", strFilename)
                    
                    'Ahora muestra al usuario un mensaje de error general
AskWhatToDo:
                    strMsg = ResolveResString(resCOMMON_CANTREG, "|1", strFilename)
                    
                    Select Case MsgError(strMsg, vbExclamation Or vbAbortRetryIgnore, gstrTitle)
                        Case vbAbort:
                            ExitSetup frmSetup1, gintRET_ABORT
                            GoTo AskWhatToDo
                        Case vbRetry:
                            GoTo RetryDllSelfReg
                        Case vbIgnore:
                            AbortAction
                        'End Case
                    End Select
                Else
                    CommitAction
                End If
            Case mstrEXESELFREGISTER
                '
                'Sólo registra automáticamente los archivos EXE
                '
                If Extension(strFilename) = strEXT_EXE Then
                    NewAction gstrKEY_EXESELFREGISTER, """" & strFilename & """"
                    Err = 0
                    ExeSelfRegister strFilename
                    If Err Then
                        AbortAction
                    Else
                        CommitAction
                    End If
                End If
            Case mstrREMOTEREGISTER
                NewAction gstrKEY_REMOTEREGISTER, """" & strFilename & """"
                Err = 0
                RemoteRegister strFilename, msRegInfo(intIdx)
                If Err Then
                    AbortAction
                Else
                    CommitAction
                End If
            Case mstrTLBREGISTER
                NewAction gstrKEY_TLBREGISTER, """" & strFilename & """"
                '
                ' Llama a la función de exportación RegisterTLB de VB5STKIT.DLL,
                ' que llama a LoadTypeLib y RegisterTypeLib.
                '
RetryTLBReg:
                If Not RegisterTLB(strFilename) Then
                    '
                    ' Falló el registro del archivo TLB.
                    '
                    LogError ResolveResString(resCOMMON_CANTREGTLB, "|1", strFilename)
TLBAskWhatToDo:
                    strMsg = ResolveResString(resCOMMON_CANTREGTLB, "|1", strFilename)
                    
                    Select Case MsgError(strMsg, vbExclamation Or vbAbortRetryIgnore, gstrTitle)
                        Case vbAbort:
                            ExitSetup frmSetup1, gintRET_ABORT
                            GoTo TLBAskWhatToDo
                        Case vbRetry:
                            GoTo RetryTLBReg
                        Case vbIgnore:
                            AbortAction
                        'End Case
                    End Select
                Else
                    CommitAction
                End If
            Case mstrVBLREGISTER
                '
                ' RegisterVBLFile se encarga del registro, etc.
                '

                RegisterVBLFile strFilename
            Case Else
                RegEdit msRegInfo(intIdx).strRegister
            'End Case
        End Select
    Next


    Erase msRegInfo

RFCleanup:
    Err = 0
End Sub
'-----------------------------------------------------------
' Procedimiento: RegisterLicenses
'
' Busca todas las entradas de licencia de setup.lst y las
' registra.
'-----------------------------------------------------------
'
Sub RegisterLicenses()
    Const strINI_LICENSES = "Licenses"
    Dim iLic As Integer
    Dim strLine As String
    Dim strLicKey As String
    Dim strLicVal As String
    Dim iCommaPos As Integer
    Dim strMsg As String

    iLic = 1
    Do
        strLine = ReadIniFile(gstrSetupInfoFile, strINI_LICENSES, gstrINI_LICENSE & iLic)
        If strLine = gstrNULL Then
            '
            ' Tenemos todas las licencias.
            '
            Exit Sub
        End If
        strLine = strUnQuoteString(strLine)
        '
        ' Tenemos una licencia; la analizamos y registramos.
        '
        iCommaPos = InStr(strLine, gstrCOMMA)
        If iCommaPos = 0 Then
            '
            ' Parece que el archivo setup.lst está dañado. En la información
              ' de licencia siempre debe haber una coma que separe la clave
            ' del valor de licencia.
            '
            GoTo RLError
        End If
        strLicKey = Left(strLine, iCommaPos - 1)
        strLicVal = Mid(strLine, iCommaPos + 1)
        
        RegisterLicense strLicKey, strLicVal
        
        iLic = iLic + 1
    Loop While strLine <> gstrNULL
    Exit Sub
        
RLError:
    strMsg = gstrSetupInfoFile & LS$ & ResolveResString(resINVLINE) & LS$
    strMsg = strMsg & ResolveResString(resSECTNAME) & strINI_LICENSES & LF$ & strLine
    MsgError strMsg, MB_ICONSTOP, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Sub
'-----------------------------------------------------------
' Procedimiento: RegisterLicense
'
' Registra información de licencia dada la clave y el valor
' predeterminado. La información de licencia siempre va en
' HKEY_CLASSES_ROOT\Licenses.
'-----------------------------------------------------------
'
Sub RegisterLicense(strLicKey As String, strLicVal As String)
    Const strREG_LICENSES = "Licenses"
    Dim hKey As Long
    '
    ' Crea la clave
    '
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strREG_LICENSES, strLicKey, hKey) Then
        '
        ' RegCreateKey muestra un error si algo va mal.
        '
        GoTo REGError
    End If
    '
    ' Establece el valor de la clave
    '
    If Not RegSetStringValue(hKey, "", strLicVal, True) Then
        '
        ' RegSetStringValue muestra un error si algo va mal.
        '
        GoTo REGError
    End If
    '
    ' Cierra la clave
    '
    RegCloseKey hKey

    Exit Sub
        
REGError:
    '
    ' Ya se deben haber mostrado mensajes de error.
    '
    ExitSetup frmSetup1, gintRET_FATAL
End Sub
'-----------------------------------------------------------
' Procedimiento: RegisterVBLFile
'
' Registra información de licencia en un archivo de licencia
' de VB (vbl).
' Básicamente, analiza la información de licencia y llama a
' RegisterLicense.
'
' Si strVBLFile no es un archivo VBL válido, no se registra
' nada.
'-----------------------------------------------------------
'
Sub RegisterVBLFile(strVBLFile As String)
    Dim strLicKey As String
    Dim strLicVal As String
    
    GetLicInfoFromVBL strVBLFile, strLicKey, strLicVal
    If strLicKey <> gstrNULL Then
        RegisterLicense strLicKey, strLicVal
    End If
End Sub

'----------------------------------------------------------
' Procedimiento: RegisterAppRemovalEXE
'
' Registra el programa para quitar aplicaciones (sólo en Windows 95)
' o coloca un icono para el mismo en el directorio de la aplicación.
'
' Devuelve True si tiene éxito; de lo contrario, devuelve False.
'----------------------------------------------------------
Function RegisterAppRemovalEXE(ByVal strAppRemovalEXE As String, ByVal strAppRemovalLog As String, ByVal strGroupName As String) As Boolean
    On Error GoTo Err
    
    Const strREGSTR_VAL_AppRemoval_APPNAMELINE = "ApplicationName"
    Const strREGSTR_VAL_AppRemoval_DISPLAYNAME = "DisplayName"
    Const strREGSTR_VAL_AppRemoval_COMMANDLINE = "UninstallString"
    Const strREGSTR_VAL_AppRemoval_APPTOUNINSTALL = "AppToUninstall"
    
    
    Dim strREGSTR_PATH_UNINSTALL As String
    strREGSTR_PATH_UNINSTALL = RegPathWinCurrentVersion() & "\Uninstall"
    
    'La línea de comandos para el ejecutable para quitar aplicaciones es
    'simplemente la ruta de acceso para el archivo de registro de la
    'instalación
    Dim strAppRemovalCmdLine As String
    strAppRemovalCmdLine = GetAppRemovalCmdLine(strAppRemovalEXE, strAppRemovalLog, gstrNULL, False, APPREMERR_NONE)
    '
    ' Se asegura de que la línea de comandos para quitar aplicaciones
    ' (incluyendo la ruta de acceso, el nombre de archivo, los
    ' argumentos, etc.) no es mayor que el máximo permitido,
    ' que es _MAX_PATH.
    '
    If Not fCheckFNLength(strAppRemovalCmdLine) Then
        Dim strMsg As String
        strMsg = ResolveResString(resCANTCREATEICONPATHTOOLONG) & LS$ & ResolveResString(resCHOOSENEWDEST) & LS$ & strAppRemovalCmdLine
        Call MsgError(strMsg, vbOKOnly, gstrSETMSG)
        ExitSetup frmCopy, gintRET_FATAL
        Exit Function
    End If
    '
    ' Crea entradas del registro para indicar a Windows dónde está el ejecutable
    ' para quitar aplicaciones, cómo debe mostrarse al usuario y cuáles son los
    ' argumentos de la línea de comandos
    '
    Dim iAppend As Integer
    Dim fOk As Boolean
    Dim hkeyAppRemoval As Long
    Dim hkeyOurs As Long
    Dim i As Integer
    
    'Sigue y crea una clave para la rama principal Uninstall
    If Not RegCreateKey(HKEY_LOCAL_MACHINE, strREGSTR_PATH_UNINSTALL, "", hkeyAppRemoval) Then
        GoTo Err
    End If
    
    'Necesitamos una clave única. Esta clave no se muestra nunca al usuario final.
    'Usaremos una clave con el formato 'ST5UNST #xxx'
    Dim strAppRemovalKey As String
    Dim strAppRemovalKeyBase As String
    Dim hkeyTest As Long
    strAppRemovalKeyBase = mstrFILE_APPREMOVALLOGBASE$ & " #"
    iAppend = 1
    
    Do
        strAppRemovalKey = strAppRemovalKeyBase & Format(iAppend)
        If RegOpenKey(hkeyAppRemoval, strAppRemovalKey, hkeyTest) Then
            'Esta clave ya existe. Pero necesitamos una clave única.
            RegCloseKey hkeyTest
        Else
            'Hemos encontrado una clave que no existe. La usamos.
            Exit Do
        End If
        
        iAppend = iAppend + 1
    Loop
    
    '
    ' También necesitamos un nombre de presentación único. Este nombre
    ' es el único medio que tiene el usuario para identificar la
    ' aplicación que se va a quitar
    '
    Dim strDisplayName As String
    strDisplayName = gstrAppName 'Primero el nombre de la aplicación
    If Not IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
        'Segundo intento... Se agrega la ruta
        strDisplayName = strDisplayName & " (" & gstrDestDir & ")"
        If Not IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
            'Intentos posteriores... Se anexa un entero único
            Dim strDisplayNameBase As String
            
            strDisplayNameBase = strDisplayName
            iAppend = 3
            Do
                strDisplayName = strDisplayNameBase & " #" & Format(iAppend)
                If IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
                    Exit Do
                Else
                    iAppend = iAppend + 1
                End If
            Loop
        End If
    End If
    
    'Sigue y llena entradas para el ejecutable para quitar aplicaciones
    If Not RegCreateKey(hkeyAppRemoval, strAppRemovalKey, "", hkeyOurs) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_APPNAMELINE, gstrAppExe, False) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_DISPLAYNAME, strDisplayName, False) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_COMMANDLINE, strAppRemovalCmdLine, False) Then
        GoTo Err
    End If
        
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_APPTOUNINSTALL, gstrAppToUninstall, False) Then
        GoTo Err
    End If
    If Not TreatAsWin95() Then
        '
        ' En NT3.51, simplemente ponemos un icono en el EXE para quitar aplicaciones
        ' del Administrador de programas
        '
        If fMainGroupWasCreated Then
            CreateProgManItem frmSetup1, strGroupName, strAppRemovalCmdLine, ResolveResString(resAPPREMOVALICONNAME, "|1", gstrAppName)
        Else
            'Si ve este mensaje, significa que no ha personalizado correctamente Form_Load().
            'Con 32 bits y NT 3.51 siempre es necesario crear un grupo de Administrador de programas.
            MsgError ResolveResString(resNOFOLDERFORICON, "|1", strAppRemovalEXE), MB_OK Or MB_ICONEXCLAMATION, gstrTitle
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
    
    RegCloseKey hkeyAppRemoval
    RegCloseKey hkeyOurs
    
    RegisterAppRemovalEXE = True
    Exit Function
    
Err:
    If hkeyOurs Then
        RegCloseKey hkeyOurs
        RegDeleteKey hkeyAppRemoval, strAppRemovalKey
    End If
    If hkeyAppRemoval Then
        RegCloseKey hkeyAppRemoval
    End If
    
    RegisterAppRemovalEXE = False
    Exit Function
End Function

'-----------------------------------------------------------
' Función: RegOpenKey
'
' Abre una clave existente en el registro del sistema.
'
' Retorno: True si la clave se abrió correctamente; de lo
'   contrario, False
'   Si tiene éxito, phkResult se establece al controlador
'   de la clave.
'-----------------------------------------------------------
'
Function RegOpenKey(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0

    strHkey = strGetHKEYString(hKey)

    lResult = OSRegOpenKey(hKey, lpszSubKey, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegOpenKey = True
        AddHkeyToCache phkResult, strHkey & "\" & lpszSubKey
    Else
        RegOpenKey = False
    End If
End Function
'----------------------------------------------------------
' Función: RegPathWinPrograms
'
' Devuelve el nombre de la clave del registro
' "\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'----------------------------------------------------------
Function RegPathWinPrograms() As String
    RegPathWinPrograms = RegPathWinCurrentVersion() & "\Explorer\Shell Folders"
End Function
 
'----------------------------------------------------------
' Función: RegPathWinCurrentVersion
'
' Devuelve el nombre de la clave del registro
' "\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion"
'----------------------------------------------------------
Function RegPathWinCurrentVersion() As String
    RegPathWinCurrentVersion = "SOFTWARE\Microsoft\Windows\CurrentVersion"
End Function

'----------------------------------------------------------
' Función: RegQueryIntValue
'
' Recupera los datos enteros para un valor con nombre
' (strValueName = nombre) o sin nombre (strValueName = "")
' dentro de una clave del registro. Si el valor con nombre
' existe, pero sus datos no son de tipo REG_DWORD, esta
' función fallará.
'
' NOTA: no hay ninguna versión de 16 bits de esta función.
'
' Retorno: True si tiene éxito; de lo contrario, False.
'   Si tiene éxito, lData se establece al valor de datos
'   numéricos
'
'----------------------------------------------------------
Function RegQueryNumericValue(ByVal hKey As Long, ByVal strValueName As String, lData As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    
    RegQueryNumericValue = False
    
    On Error GoTo 0
    
    ' Obtiene el tipo y longitud de los datos
    lDataBufSize = 4
        
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            lData = lBuf
            RegQueryNumericValue = True
        End If
    End If
End Function

' Función: RegQueryStringValue
'
' Recupera los datos de cadena para un valor con nombre
' (strValueName = nombre) o sin nombre (strValueName = "")
' dentro de una clave del registro. Si el valor con
' nombre existe, pero sus datos no son una cadena, esta
' función fallará.
'
' NOTA: para 16 bits, strValueName DEBE ser "" (pero el
' NOTA: parámetro se deja por compatibilidad con el
' NOTA: código de origen)
'
' Retorno: True si tiene éxito; de lo contrario, False.
'   Si tiene éxito, strData se establece al valor de
'   datos de cadena
'
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    RegQueryStringValue = False
    On Error GoTo 0
    ' Obtiene el tipo y longitud de los datos
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StripTerminator(strBuf)
            End If
        End If
    End If
End Function

'----------------------------------------------------------
' Función: RegQueryRefCount
'
' Recupera los datos como número de referencias para un
' determinado valor dentro de una clave del registro.
' Aunque REG_DWORD es la forma preferida de almacenar
' números de referencia, es posible que algunos programas
' de instalación utilicen incorrectamente en su lugar un
' valor de cadena o binario. Esta rutina acepta los
' datos, ya se trate de una cadena, un valor binario o
' DWORD (Long).
'
' NOTA: no hay ninguna versión de 16 bits de esta función.
'
' Retorno: True si tiene éxito; de lo contrario, False.
'   Si tiene éxito, lrefcount se establece al valor de
'   datos numéricos
'
'----------------------------------------------------------
Function RegQueryRefCount(ByVal hKey As Long, ByVal strValueName As String, lRefCount As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long

    RegQueryRefCount = False

    On Error GoTo 0

    ' Obtiene el tipo y longitud de los datos
    lDataBufSize = 4

    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        Select Case lValueType
            Case REG_DWORD
                lRefCount = lBuf
                RegQueryRefCount = True
            Case REG_BINARY
                If lDataBufSize = 4 Then
                    lRefCount = lBuf
                    RegQueryRefCount = True
                End If
            Case REG_SZ
                Dim strRefCount As String
                
                If RegQueryStringValue(hKey, strValueName, strRefCount) Then
                    lRefCount = Val(strRefCount)
                    RegQueryRefCount = True
                End If
            'End Case
        End Select
    End If
End Function

' Función: RegSetNumericValue
'
' Asocia un valor con nombre (strValueName = nombre) o sin nombre (strValueName = "")
'   con una clave del registro.
'
' Si fLog falta o es True, esta acción se registra en el archivo de registro
' y el programa para quitar aplicaciones eliminará el valor si el usuario
' elige quitar la aplicación instalada.
'
' NOTA: no hay ninguna versión de 16 bits de esta función.
'
' Retorno: True si tiene éxito; de lo contrario, False.
'
Function RegSetNumericValue(ByVal hKey As Long, ByVal strValueName As String, ByVal lData As Long, Optional ByVal fLog) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0
    
    If IsMissing(fLog) Then fLog = True

    strHkey = strGetHKEYString(hKey)
    
    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_DWORD, lData, 4)
    If lResult = ERROR_SUCCESS Then
        RegSetNumericValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetNumericValue = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

' Función: RegSetStringValue
'
' Asocia un valor con nombre (strValueName = nombre) o sin nombre (strValueName = "")
'   con una clave del registro.
'
' Si fLog falta o es True, esta acción se registra en el archivo de registro
' y el programa para quitar aplicaciones eliminará el valor si el usuario
' elige quitar la aplicación instalada.
'
' Retorno: True si tiene éxito; de lo contrario, False.
'
Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String, Optional ByVal fLog) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    
    On Error GoTo 0
    
    If IsMissing(fLog) Then fLog = True

    If hKey = 0 Then
        Exit Function
    End If
    
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
    
    If lResult = ERROR_SUCCESS Then
        RegSetStringValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetStringValue = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

'-----------------------------------------------------------
' Procedimiento: RemoteRegister
'
' Ejecuta de forma síncrona el programa de registro del
' cliente en el archivo de registro del servidor remoto
' especificado para configurarlo correctamente en el
' registro.
'
' ENTRADA: [strFileName] - archivo .EXE que se va a registrar

'-----------------------------------------------------------
'
Sub RemoteRegister(ByVal strFilename As String, rInfo As REGINFO)
    Const strClientRegistrationUtility$ = "CLIREG32.EXE"
    Const strAddressSwitch = " /s "
    Const strProtocolSwitch = " /p "
    Const strSilentSwitch = " /q "
    Const strNoLogoSwitch = " /nologo "
    Const strAuthenticationSwitch = " /a "
    Const strTypelibSwitch = " /t "
    Const strDCOMSwitch = " /d "
    Const strEXT_REMOTE$ = "VBR"
    Const strEXT_REMOTETLB$ = "TLB"

    Dim strAddress As String
    Dim strProtocol As String
    Dim intAuthentication As Integer
    Dim strCmdLine As String
    Dim fShell As Integer
    Dim strMatchingTLB As String
    Dim fDCOM As Boolean

    'Busca el nombre del correspondiente archivo de biblioteca de tipos.
    'Ya debe haberse instalado en el mismo directorio que el archivo .VBR.
    strMatchingTLB = strFilename
    If Right$(strMatchingTLB, Len(strEXT_REMOTE)) = strEXT_REMOTE Then
        strMatchingTLB = Left$(strMatchingTLB, Len(strMatchingTLB) - Len(strEXT_REMOTE))
    End If
    strMatchingTLB = strMatchingTLB & strEXT_REMOTETLB

    strAddress = rInfo.strNetworkAddress
    strProtocol = rInfo.strNetworkProtocol
    intAuthentication = rInfo.intAuthentication
    fDCOM = rInfo.fDCOM
    frmRemoteServerDetails.GetServerDetails strFilename, strAddress, strProtocol, fDCOM
    frmMessage.Refresh
    strCmdLine = _
      strClientRegistrationUtility _
      & strAddressSwitch & """" & strAddress & """" _
      & IIf(fDCOM, " ", strProtocolSwitch & strProtocol) _
      & IIf(fDCOM, " ", strAuthenticationSwitch & Format$(intAuthentication) & " ") _
      & strNoLogoSwitch _
      & strTypelibSwitch & """" & strMatchingTLB & """" & " " _
      & IIf(fDCOM, strDCOMSwitch, "") _
      & IIf(gfNoUserInput, strSilentSwitch, "") _
      & """" & strFilename & """"
      
    '
    'Sale de forma síncrona y ejecuta el programa con los modificadores correctos
    '
    fShell = FSyncShell(strCmdLine, vbNormal)

    If Not fShell Then
        MsgError ResolveResString(resCANTRUNPROGRAM, "|1", strClientRegistrationUtility), vbOKOnly Or vbExclamation, gstrTitle, gintRET_FATAL
        ExitSetup frmSetup1, gintRET_FATAL
    End If
End Sub

'-----------------------------------------------------------
' Procedimiento: RemoveShellLink
'
' Quita un vínculo de Inicio>Programas o cualquiera de sus
' subcarpetas inmediatas en el intérprete de comandos de
' Windows 95.
'
' ENTRADA: [strFolderName] - nombre de texto de la carpeta
'                       inmediata en la que existe
'                       actualmente el vínculo que se va
'                       a quitar o bien la cadena vacía
'                       ("") para indicar que el vínculo
'                       se encuentra directamente en el
'                       menú Inicio>Programas.
'     [strLinkName] - título de texto para el vínculo
'
' Esta acción nunca se registra en el archivo de registro
' para quitar aplicaciones.
'
' CONDICIÓN PREVIA: strFolderName ya se ha creado y es una
'               subcarpeta inmediata de Inicio>Programas,
'               si no es igual a ""
'-----------------------------------------------------------
'
Sub RemoveShellLink(ByVal strFolderName As String, ByVal strLinkName As String)
    Dim fSuccess As Boolean
    
    ReplaceDoubleQuotes strFolderName
    ReplaceDoubleQuotes strLinkName
    
    fSuccess = OSfRemoveShellLink(strFolderName, strLinkName)
End Sub

'-----------------------------------------------------------
' Función: ResolveDestDir
'
' Dada una cadena de directorio de destino, hace equivaler
' cualquier fragmento de macro de la cadena con sus
' ubicaciones reales determinadas en tiempo de ejecución
' y devuelve una cadena que refleja la ruta de acceso real.
'
' ENTRADA: [strDestDir] - cadena que contiene información de
'                    macro de directorio y de la ruta real
'                    del directorio
'
'     [fAssumeDir] - boolean si es True; hace que esta rutina
'                    suponga que strDestDir contiene una ruta
'                    de acceso de directorio. Si no se indica
'                    ningún directorio, lo hará la ruta de
'                    la aplicación. Si es False, esta rutina
'                    devolverá strDestDir después de realizar
'                    la expansión. Establézcalo a False cuando
'                    no sepa con seguridad si un directorio
'                    pero desee expandir macros, si contiene
'                    alguna. Por ejemplo, si es un parámetro
'                    de la línea de comandos, no podrá saber
'                    con seguridad si se refiere a una ruta
'                    de acceso. En este caso, establezca
'                    fAssumeDir = False. El valor predeterminado
'                    es True.
'
' Retorno: una cadena que contiene el nombre de directorio
' resuelto
'-----------------------------------------------------------
'
Function ResolveDestDir(ByVal strDestDir As String, Optional fAssumeDir As Variant) As String
    Const strMACROSTART$ = "$("
    Const strMACROEND$ = ")"

    Dim intPos As Integer
    Dim strResolved As String
    Dim hKey As Long
    Dim strPathsKey As String
    Dim fQuoted As Boolean
    
    If IsMissing(fAssumeDir) Then
        fAssumeDir = True
    End If
    
    strPathsKey = RegPathWinCurrentVersion()
    strDestDir = Trim(strDestDir)
    '
    ' Si strDestDir está entre comillas cuando se pasa a esta
    ' rutina, debe aparecer entre comillas cuando se devuelva.
    ' Sin embargo, hay que quitar temporalmente las comillas
    ' para su proceso.
    '
    If Left(strDestDir, 1) = gstrQUOTE Then
        fQuoted = True
        strDestDir = strUnQuoteString(strDestDir)
    End If
    '
    ' Tomamos la primera parte del directorio de destino y, si es $(, necesitamos
    ' la parte que va hasta el último paréntesis (incluido). Después lo probamos
    ' de nuevo para la expansión de macros. Si no se encuentra ) después de $(,
    ' debemos suponer que simplemente se trata de un nombre de archivo normal y
    ' no se procesa. Incluya únicamente la instrucción Case si strDestDir empieza
    ' con $(.
    '
    If Left$(strDestDir, 2) = strMACROSTART Then
        intPos = InStr(strDestDir, strMACROEND)

        Select Case Left$(strDestDir, intPos)
            Case gstrAPPDEST
                If gstrDestDir <> gstrNULL Then

                    strResolved = gstrDestDir
                Else
                    strResolved = "?"
                End If
            Case gstrWINDEST
                strResolved = gstrWinDir
            Case gstrWINSYSDEST, gstrWINSYSDESTSYSFILE
                strResolved = gstrWinSysDir
            Case gstrPROGRAMFILES
                If TreatAsWin95() Then
                    Const strProgramFilesKey = "ProgramFilesDir"
    
                    If RegOpenKey(HKEY_LOCAL_MACHINE, strPathsKey, hKey) Then
                        RegQueryStringValue hKey, strProgramFilesKey, strResolved
                        RegCloseKey hKey
                    End If
                End If
    
                If strResolved = "" Then
                    'Si no se establece de otra forma, deja que strResolved
                          'sea la raíz del primer disco fijo
                    strResolved = strRootDrive()
                End If
            Case gstrCOMMONFILES
                'Primero determina la ruta correcta de Archivos de programa\Archivos
                    'comunes, si estamos bajo Win95
                strResolved = strGetCommonFilesPath()
                If strResolved = "" Then
                    'Si no se establece de otra forma, deja que strResolved sea
                          'el directorio Windows
                    strResolved = gstrWinDir
                End If
            Case gstrCOMMONFILESSYS
                'Primero determina la ruta correcta de Archivos de programa\Archivos
                    'comunes, si estamos bajo Win95
                Dim strCommonFiles As String
                
                strCommonFiles = strGetCommonFilesPath()
                If strCommonFiles <> "" Then
                    'Ahora agregamos \System y ya terminamos
                    strResolved = strCommonFiles & "System\"
                Else
                    'Si Archivos comunes no está en el registro, asigna toda la
                    'macro al directorio Windows\{system,system32}
                    strResolved = gstrWinSysDir
                End If
            Case gstrDAODEST
                strResolved = strGetDAOPath()
            Case Else
                intPos = 0
            'End Case
        End Select
    End If
    
    If intPos <> 0 Then
        AddDirSep strResolved
    End If

    If fAssumeDir = True Then
        If intPos = 0 Then
            '
            'si no hay especificación de unidad y no empieza con cualquier indicador
                'de ruta raíz ("\"), suponemos que este destino es relativo al directorio
                'de destino de la aplicación
            '
            If Mid$(strDestDir, 2, 1) <> gstrCOLON Then
                If Left$(strDestDir, 1) <> gstrSEP_DIR Then
                    strResolved = gstrDestDir
                End If
            End If
        Else
            If Mid$(strDestDir, intPos + 1, 1) = gstrSEP_DIR Then
                intPos = intPos + 1
            End If
        End If
    End If

    If fQuoted = True Then
        ResolveDestDir = strQuoteString(strResolved & Mid$(strDestDir, intPos + 1), True, False)
    Else
        ResolveDestDir = strResolved & Mid$(strDestDir, intPos + 1)
    End If
End Function
'-----------------------------------------------------------
' Función: ResolveDestDirs
'
' Dada una cadena delimitada por espacios, esta rutina busca
' todas las macros de directorio de destino y las expande
' haciendo repetidas llamadas a ResolveDestDir. Vea
' ResolveDestDir.
'
' Observe que la macro debe ir inmediatamente a continuación
' de un delimitador de espacio en blanco (o un espacio en
' blanco seguido de una comilla); de lo contrario, se
' pasará por alto.
'
' Esta rutina no supone que cada elemento de la cadena
' delimitada sea realmente una ruta de directorio. Por
' tanto, el último parámetro de la llamada a ResolveDestDir,
' a continuación, es False.
'
' ENTRADA: [str] - cadena que contiene información de macro o
'             macros de directorio y de la ruta real del
'             directorio
'
' Retorno: cadena con macros de directorio de destino
' expandidas.
'-----------------------------------------------------------
'
Function ResolveDestDirs(str As String)
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim strField As String
    Dim strExpField As String
    Dim strExpanded As String
    
    If Len(Trim(strUnQuoteString(str))) = 0 Then
        ResolveDestDirs = str
        Exit Function
    End If
        
    intAnchor = 1
    strExpanded = ""
    
    Do
        intOffset = intGetNextFldOffset(intAnchor, str, " ")
        If intOffset = 0 Then intOffset = Len(str) + 1
        strField = Mid(str, intAnchor, intOffset - intAnchor)
        strExpField = ResolveDestDir(strField, False)
        strExpanded = strExpanded & strExpField & " "
        intAnchor = intOffset + 1
    Loop While intAnchor < Len(str)
    
    ResolveDestDirs = Trim(strExpanded)
End Function
'-----------------------------------------------------------
' Función: ResolveDir
'
' Dado un nombre de ruta de acceso, se resuelve a su forma
' más corta. Si la ruta no es válida, se avisa opcionalmente
' al usuario.
'
' ENTRADA: [strPathName] - ruta de acceso que se va a resolver
'     [fMustExist] - exige que la ruta exista realmente
'     [fWarn] - si es True, avisa al usuario si la ruta no
'               es válida
'
' Retorno: cadena que contiene el nombre de directorio resuelto
'-----------------------------------------------------------
'
Function ResolveDir(ByVal strPathName As String, fMustExist As Integer, fWarn As Integer) As String
    Dim strMsg As String
    Dim fInValid As Integer
    Dim strUnResolvedPath As String
    Dim strResolvedPath As String
    Dim strIgnore As String
    Dim cbResolved As Long

    On Error Resume Next

    fInValid = False
    '
    'Si la ruta de acceso es un nombre UNC (sólo para 16 bits), o si está
    'en un nombre de archivo, entonces no es válida
    '
    If FileExists(strPathName) = True Then
        fInValid = True
        GoTo RDContinue
    End If

    strUnResolvedPath = strPathName

    If InStr(3, strUnResolvedPath, gstrSEP_DIR) > 0 Then

        strResolvedPath = Space(gintMAX_PATH_LEN * 2)
        cbResolved = GetFullPathName(strUnResolvedPath, gintMAX_PATH_LEN, strResolvedPath, strIgnore)
        If cbResolved = 0 Then
            '
            ' No se pudo resolver la ruta de acceso. Si podemos cambiar
                ' realmente al directorio deseado, continuamos de todos modos.
            '
            ChDir strUnResolvedPath
            AddDirSep strUnResolvedPath
            If Err > 0 Then
                Err = 0
                ChDir strUnResolvedPath
                If Err > 0 Then
                    fInValid = True
                Else
                    strResolvedPath = strUnResolvedPath
                End If
            Else
                strResolvedPath = strUnResolvedPath
            End If
        Else
            '
            ' GetFullPathName devolvió una cadena terminada en NULL
            ' en strResolvedPath. Quita el NULL.
            '
            strResolvedPath = StripTerminator(strResolvedPath)
            If CheckDrive(strResolvedPath, gstrTitle) = False Then
                fInValid = True
            Else
                AddDirSep strResolvedPath
                If fMustExist = True Then
                    Err = 0
                    
                    Dim strDummy As String
                    strDummy = Dir$(strResolvedPath & "*.*")
                    
                    If Err > 0 Then
                        strMsg = ResolveResString(resNOTEXIST) & LS$
                        fInValid = True
                    End If
                End If
            End If
        End If
    Else
        fInValid = True
    End If

RDContinue:
    If fInValid = True Then
        If fWarn = True Then
            strMsg = strMsg & ResolveResString(resDIRSPECIFIED) & LS$ & strPathName & LS$
            strMsg = strMsg & ResolveResString(resDIRINVALID)
            MsgError strMsg, MB_OK Or MB_ICONEXCLAMATION, ResolveResString(resDIRINVNAME)
            If gfNoUserInput Then
                ExitSetup frmSetup1, gintRET_FATAL
            End If
        End If

        ResolveDir = gstrNULL
    Else
        ResolveDir = strResolvedPath
    End If

    Err = 0
End Function

'-----------------------------------------------------------
' Procedimiento: RestoreProgMan
'
' Restaura el Administrador de programas de Windows
'-----------------------------------------------------------
'
Sub RestoreProgMan()
    Const strPMTITLE$ = "Program Manager"

    On Error Resume Next

    'Prueba primero con el nombre traducido
    AppActivate ResolveResString(resPROGRAMMANAGER)
    
    If Err Then
        'Si no funciona, prueba con el nombre inglés
        AppActivate strPMTITLE
    End If

    Err = 0
End Sub

'-----------------------------------------------------------
' Función: SetFileDateTime
'
' Establece la fecha y hora del archivo de destino a la
' fecha y hora del archivo de origen
'
' ENTRADA: [strFileGetTime] - archivo del que obtener la
'                             información de fecha y hora
'     [strFileSetTime] - archivo para el que se va a
'                        establecer la información de fecha
'                        y hora
'
' Retorno: True si la fecha y la hora se han establecido
' correctamente; de lo contrario es False
'-----------------------------------------------------------
'
Function SetFileDateTime(strFileGetTime As String, strFileSetTime As String) As Integer
    SetFileDateTime = IIf(SetTime(strFileGetTime, strFileSetTime) = -1, False, True)
End Function

'-----------------------------------------------------------
' Procedimiento: ShowPathDialog
'
' Muestra un formulario que permite al usuario obtener una
' ruta de origen o de destino
'
' ENTRADA: [strPathRequest] - determina si se pedirá el
'                        nombre de ruta de origen o de destino.
'                        gstrDIR_SRC para ruta de origen
'                        gstrDIR_DEST para ruta de destino
'-----------------------------------------------------------
'
Sub ShowPathDialog(ByVal strPathRequest As String)
    frmSetup1.Tag = strPathRequest

    '
    'frmPath.Form_Load() lee frmSetup1.Tag para determinar si
    'es una petición de la ruta de origen o la de destino
    '
    frmPath.Show 1

    If strPathRequest = gstrDIR_SRC Then
        gstrSrcPath = frmSetup1.Tag
    Else
        If gfRetVal = gintRET_CONT Then
            gstrDestDir = frmSetup1.Tag
        End If
    End If
End Sub

'-----------------------------------------------------------
' Función: strExtractFilenameArg
'
' Extrae un nombre de archivo entre comillado o sin comillas
'   de una cadena que contiene argumentos de la línea
'   de comandos
'
' ENTRADA: [str] - cadena que contiene un nombre de archivo.
'             Este nombre empieza en el primer carácter y
'             continúa hasta el final de la cadena, hasta
'             el primer espacio en blanco o carácter
'             modificador o, si la cadena empieza con
'             comillas dobles, continúa hasta la siguiente
'             comilla doble
' SALIDA: devuelve el nombre de archivo, sin comillas
'      str se establece para que sea el resto de la cadena
'      después del nombre de archivo y la comilla (si hay alguna)
'
'-----------------------------------------------------------
'
Function strExtractFilenameArg(str As String, fErr As Boolean)
    Dim strFilename As String
    
    str = Trim$(str)
    
    Dim iEndFilenamePos As Integer
    If Left$(str, 1) = """" Then
        ' Los nombres de archivo van entrecomillados
        iEndFilenamePos = InStr(2, str, """") ' Busca comilla coincidente
        If iEndFilenamePos > 0 Then
            strFilename = Mid$(str, 2, iEndFilenamePos - 2)
            str = Right$(str, Len(str) - iEndFilenamePos)
        Else
            fErr = True
            Exit Function
        End If
    Else
        ' El nombre de archivo continúa hasta el siguiente modificador,
        ' espacio en blanco o comilla
        Dim iSpacePos As Integer
        Dim iSwitch1 As Integer
        Dim iSwitch2 As Integer
        Dim iQuote As Integer
        
        iSpacePos = InStr(str, " ")
        iSwitch1 = InStr(str, gstrSwitchPrefix1)
        iSwitch2 = InStr(str, gstrSwitchPrefix2)
        iQuote = InStr(str, """")
        
        If iSpacePos = 0 Then iSpacePos = Len(str) + 1
        If iSwitch1 = 0 Then iSwitch1 = Len(str) + 1
        If iSwitch2 = 0 Then iSwitch2 = Len(str) + 1
        If iQuote = 0 Then iQuote = Len(str) + 1
        
        iEndFilenamePos = iSpacePos
        If iSwitch1 < iEndFilenamePos Then iEndFilenamePos = iSwitch1
        If iSwitch2 < iEndFilenamePos Then iEndFilenamePos = iSwitch2
        If iQuote < iEndFilenamePos Then iEndFilenamePos = iQuote
        
        strFilename = Left$(str, iEndFilenamePos - 1)
        If iEndFilenamePos > Len(str) Then
            str = ""
        Else
            str = Right(str, Len(str) - iEndFilenamePos + 1)
        End If
    End If
    
    strFilename = Trim$(strFilename)
    If strFilename = "" Then
        fErr = True
        Exit Function
    End If
    
    fErr = False
    strExtractFilenameArg = strFilename
    str = Trim$(str)
End Function



'-----------------------------------------------------------
' Procedimiento: UpdateStatus
'
' "Llena" (según un porcentaje) el interior del PictureBox y
' muestra el porcentaje relleno
'
' ENTRADA: [pic] - PictureBox usado para la región de "relleno"
'     [sngPercent] - Porcentaje de la forma que se rellenará
'     [fBorderCase] - Indica si el porcentaje especificado
'        incluye o no al borde; es decir, exactamente el 0%
'        o exactamente el 100%. A menos que fBorderCase sea
'        True, se supondrá que los valores 0% y 100% serán
'        "cercanos" a esos valores y en su lugar se usarán
'        1% y 99%.
'
' Notas: establezca a True la propiedad AutoRedraw del PictureBox
'        de forma que se puedan volver a dibujar automáticamente
'        la barra de estado y el porcentaje, si es necesario
'-----------------------------------------------------------
'
Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    If IsMissing(fBorderCase) Then fBorderCase = False
    
    'Para que esto funcione, necesitamos un fondo blanco y un primer plano
    'de cualquier color (azul)
    Const colBackground = &HFFFFFF ' blanco
    Const colForeground = &H800000 ' azul oscuro

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Da formato al porcentaje y obtiene atributos de texto
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Nunca permite que el porcentaje sea 0 ó 100, a menos que sea exactamente
    'ese valor. Esto impide, por ejemplo, que la barra de estado sea del 100%
    'hasta que hayamos terminado completamente.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Ahora se establece intX e intY a la ubicación inicial para imprimir el porcentaje
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Necesitamos dibujar un cuadro relleno con el color de fondo para
    'evitar el porcentaje anterior (si hay alguno)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Vuelve a la posición central de impresión e imprime el texto
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Ahora llena el cuadro con el color de cinta hasta el porcentaje deseado
    'Si el porcentaje es 0, llena todo el cuadro con el color de fondo para limpiarlo
    'Usa el valor "Not XOR" para cambiar el color del texto a blanco siempre que lo
    'toquemos y para cambiar el color del fondo a azul siempre que lo toquemos.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
End Sub

'-----------------------------------------------------------
' Función: WriteAccess
'
' Determina si hay acceso de escritura al directorio
' especificado.
'
' ENTRADA: [strDirName] - directorio para el que se va a
'                         comprobar si hay acceso de escritura
'
' Retorno: True si hay acceso de escritura; de lo contrario, False
'-----------------------------------------------------------
'
Function WriteAccess(ByVal strDirName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    AddDirSep strDirName

    intFileNum = FreeFile
    Open strDirName & mstrCONCATFILE For Output As intFileNum

    WriteAccess = IIf(Err, False, True)
    
    Close intFileNum

    Kill strDirName & mstrCONCATFILE

    Err = 0
End Function
'-----------------------------------------------------------
' Función: WriteMIF
'
' Si es una instalación SMS, esta rutina escribe el archivo
' de estado MIF que falló si algo va mal o de MIF correcto
' si todo se instala correctamente.
'
' El archivo MIF requiere un formato especial especificado
' por SMS. Actualmente, esta rutina implementa los requisitos
' mínimos. Las siguientes cadenas no modificadas que se
' escriben en MIF se deben escribir carácter a carácter tal
' y como están; excepto que el mensaje de estado debe cambiar
' dependiendo de las circunstancias de la instalación. NO
' TRADUCIR nada excepto el mensaje de estado.
'
' ENTRADA: [strMIFFilename] - el nombre del archivo MIF.
'                        setup.exe lo pasa a setup1.
'                        Probablemente se llama
'                        <aplicación>.mif, donde
'                        <aplicación> es el nombre de la
'                        aplicación que va a instalar.
'
'     [fStatus] - False si se escribe en un MIF fallido
'                 (la instalación falló); True si se escribe
'                 en un MIF correctamente.
'
'     [strSMSDescription] - es la cadena de descripción que se
'                           escribirá en el archivo MIF. No
'                           puede tener más de 255 caracteres
'                           y no puede contener retornos de
'                           carro y avances de línea. Esta
'                           rutina exigirá estos requisitos.
'
' Cuando esté en modo SMS, tenga en cuenta que no hay otra
' forma de mostrar un mensaje al usuario que no sea
' escribirlo en el archivo MIF. Mostrará un MsgBox hará
' que parezca que el equipo se ha quedado bloqueado.
' Por tanto, esta rutina no intenta mostrar ningún
' mensaje de error.
'
'-----------------------------------------------------------
'
Sub WriteMIF(ByVal strMIFFilename As String, ByVal fStatus As Boolean, ByVal strSMSDescription As String)
    Const strSUCCESS = """SUCCESS"""                 ' No se puede traducir por SMS
    Const strFAILED = """FAILED"""                   ' No se puede traducir por SMS
    
    Dim fn As Integer
    Dim intOffset As Integer
    Dim fOpened As Boolean
        
    fOpened = False
        
    On Error GoTo WMIFFAILED  ' Si falla, volvemos sin hacer nada, ya que no hay
                              ' ninguna forma de informar al usuario mientras
                              ' estamos en modo SMS.

    '
    ' Si la cadena de descripción tiene más de 255 caracteres,
    ' se trunca. Requerido por SMS.
    '
    strSMSDescription = Left(strSMSDescription, MAX_SMS_DESCRIP)
    '
    ' Quita los retornos de carro o avances de línea que pueda haber
    ' y los reemplaza por espacios en blanco. El mensaje debe
    ' tener una única línea.
    '
    For intOffset = 1 To Len(strSMSDescription)
        If (Mid(strSMSDescription, intOffset, 1) = Chr(10)) Or (Mid(strSMSDescription, intOffset, 1) = Chr(13)) Then
            Mid(strSMSDescription, intOffset, 1) = " "
        End If
    Next intOffset
    '
    ' Abre el archivo MIF en modo anexar, pero primero elimina
    ' cualquier archivo existente que tenga el mismo nombre.
    ' Observe que setup.exe pasó un nombre único; por eso, si
    ' hay alguno con ese nombre en el disco, lo puso allí
    ' setup.exe.
    '
    If FileExists(strMIFFilename) Then
        Kill strMIFFilename
    End If
    
    fn = FreeFile
    Open strMIFFilename For Append As fn
    fOpened = True
    '
    ' Estamos preparados para escribir en el archivo MIF real
    ' No hay que traducir ninguna de las cadenas siguientes.
    '
    Print #fn, "Start Component"
        Print #fn, Tab; "Name = ""Workstation"""
        Print #fn, Tab; "Start Group"
            Print #fn, Tab; Tab; "Name = ""InstallStatus"""
            Print #fn, Tab; Tab; "ID = 1"
            Print #fn, Tab; Tab; "Class = ""MICROSOFT|JOBSTATUS|1.0"""
            Print #fn, Tab; Tab; "Start Attribute"
                Print #fn, Tab; Tab; Tab; "Name = ""Status"""
                Print #fn, Tab; Tab; Tab; "ID = 1"
                Print #fn, Tab; Tab; Tab; "Type = String(16)"
                Print #fn, Tab; Tab; Tab; "Value = "; IIf(fStatus, strSUCCESS, strFAILED)
            Print #fn, Tab; Tab; "End Attribute"
            Print #fn, Tab; Tab; "Start Attribute"
                Print #fn, Tab; Tab; Tab; "Name = ""Description"""
                Print #fn, Tab; Tab; Tab; "ID = 2"
                Print #fn, Tab; Tab; Tab; "Type = String(256)"
                Print #fn, Tab; Tab; Tab; "Value = "; strSMSDescription
            Print #fn, Tab; Tab; "End Attribute"
        Print #fn, Tab; "End Group"
    Print #fn, "End Component"

    Close fn
    '
    ' Se hizo correctamente
    '
    Exit Sub

WMIFFAILED:
    '
    ' En este momento no podemos crear el archivo MIF.
    ' Como estamos en SMS, no podemos avisar a nadie,
    ' por lo que no generamos ningún mensaje de error.
    '
    If fOpened = True Then
        Close fn
    End If
    Exit Sub
End Sub

'Agrega o reemplaza una HKEY a la lista de HKEY de la memoria caché.
'Observe que esto no es necesario para quitar claves de la lista.
Private Sub AddHkeyToCache(ByVal hKey As Long, ByVal strHkey As String)
    Dim intIdx As Integer
    
    intIdx = intGetHKEYIndex(hKey)
    If intIdx < 0 Then
        'La clave no existe. Se agrega al final.
        On Error Resume Next
        ReDim Preserve hkeyCache(0 To UBound(hkeyCache) + 1)
        If Err Then
            'Si hubo un error, significa que la memoria caché estaba vacía.
            On Error GoTo 0
            ReDim hkeyCache(0 To 0)
        End If
        On Error GoTo 0

        intIdx = UBound(hkeyCache)
    Else
        'La clave ya existe. Se reemplazará.
    End If

    hkeyCache(intIdx).hKey = hKey
    hkeyCache(intIdx).strHkey = strHkey
End Sub

'Dada una HKEY predefinida, devuelve la cadena de texto que representa
'dicha clave; de lo contrario, devuelve "".
Private Function strGetPredefinedHKEYString(ByVal hKey As Long) As String
    Select Case hKey
        Case HKEY_CLASSES_ROOT
            strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            strGetPredefinedHKEYString = "HKEY_USERS"
        'End Case
    End Select
End Function

'Dada una HKEY, devuelve la cadena de texto que representa
'dicha clave.
Private Function strGetHKEYString(ByVal hKey As Long) As String
    Dim strKey As String

    '¿Está la hkey predefinida?
    strKey = strGetPredefinedHKEYString(hKey)
    If strKey <> "" Then
        strGetHKEYString = strKey
        Exit Function
    End If
    
    'No está predefinida. Buscamos en la memoria caché.
    Dim intIdx As Integer
    intIdx = intGetHKEYIndex(hKey)
    If intIdx >= 0 Then
        strGetHKEYString = hkeyCache(intIdx).strHkey
    Else
        strGetHKEYString = ""
    End If
End Function

'Busca en la memoria caché el índice de la HKEY dada.
'Devuelve el índice si se encuentra; de lo contrario devuelve -1.
Private Function intGetHKEYIndex(ByVal hKey As Long) As Integer
    Dim intUBound As Integer
    
    On Error Resume Next
    intUBound = UBound(hkeyCache)
    If Err Then
        'Si hubo un error al tener acceso al límite superior de la matriz,
        'la memoria caché está vacía
        GoTo NotFound
    End If
    On Error GoTo 0

    Dim intIdx As Integer
    For intIdx = 0 To intUBound
        If hkeyCache(intIdx).hKey = hKey Then
            intGetHKEYIndex = intIdx
            Exit Function
        End If
    Next intIdx
    
NotFound:
    intGetHKEYIndex = -1
End Function

'Devuelve la ubicación de la ruta de acceso Archivos de programa\Archivos comunes,
'si está presente en el registro. De lo contrario, devuelve "".
Public Function strGetCommonFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    If TreatAsWin95() Then
        Const strCommonFilesKey = "CommonFilesDir"

        If RegOpenKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), hKey) Then
            RegQueryStringValue hKey, strCommonFilesKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetCommonFilesPath = strPath
End Function
'Devuelve la ubicación de la ruta de acceso "Windows\Menú Inicio\Programas",
'si está presente en el registro. De lo contrario, devuelve "".
Public Function strGetProgramsFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    strPath = ""
    If TreatAsWin95() Then
        Const strProgramsKey = "Programs"

        If RegOpenKey(HKEY_CURRENT_USER, RegPathWinPrograms(), hKey) Then
            RegQueryStringValue hKey, strProgramsKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetProgramsFilesPath = strPath
End Function

'Devuelve el directorio donde está o debe estar instalado DAO. Si la clave
'no existe en el registro, se crea. Por ejemplo, en NT 3.51 esta ubicación
'suele ser 'C:\WINDOWS\MSAPPS\DAO'
Private Function strGetDAOPath() As String
    Const strMSAPPS$ = "MSAPPS\"
    Const strDAO3032$ = "DAO3032.DLL"
    
    'primero busca en el registro
    Const strKey = "SOFTWARE\Microsoft\Shared Tools\DAO"
    Const strValueName = "Path"
    Dim hKey As Long
    Dim strPath As String

    If RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey) Then
        RegQueryStringValue hKey, strValueName, strPath
        RegCloseKey hKey
    End If

    If strPath <> "" Then
        strPath = GetPathName(strPath)
        AddDirSep strPath
        strGetDAOPath = strPath
        Exit Function
    End If
    
    'No está en el registro, por lo que necesitamos decidir
    'dónde debe estar el directorio y necesitamos incluir
    'dicha ubicación en el registro.
 

    If TreatAsWin95() Then
        'Para Win95, se usa "Archivos comunes\Microsoft Shared\DAO"
        strPath = strGetCommonFilesPath() & ResolveResString(resMICROSOFTSHARED) & "DAO\"
    Else
        'De lo contrario se usa Windows\MSAPPS\DAO
        strPath = gstrWinDir & strMSAPPS & "DAO\"
    End If
    
    'Coloca esta información en el registro (observe que señalamos a la propia DAO3032.DLL,
    'no sólo al directorio)
    If RegCreateKey(HKEY_LOCAL_MACHINE, strKey, "", hKey) Then
        RegSetStringValue hKey, strValueName, strPath & strDAO3032, False
        RegCloseKey hKey
    End If

    strGetDAOPath = strPath
End Function

' Reemplaza todas las comillas dobles por comillas simples
Public Sub ReplaceDoubleQuotes(str As String)
    Dim i As Integer
    
    For i = 1 To Len(str)
        If Mid$(str, i, 1) = """" Then
            Mid$(str, i, 1) = "'"
        End If
    Next i
End Sub

'Obtiene la parte de ruta de acceso de un nombre de archivo
Function GetPathName(ByVal strFilename As String) As String
    Dim intPos As Integer
    Dim strPathOnly As String
    Dim dirTmp As DirListBox
    Dim i As Integer

    On Error Resume Next


    Err = 0
    
    intPos = Len(strFilename)

    '
    'Cambia todos los caracteres '/' a '\'
    '

    For i = 1 To Len(strFilename)
        If Mid$(strFilename, i, 1) = gstrSEP_DIRALT Then
            Mid$(strFilename, i, 1) = gstrSEP_DIR
        End If
    Next i

    If InStr(strFilename, gstrSEP_DIR) = intPos Then
        If intPos > 1 Then
            intPos = intPos - 1
        End If
    Else
        Do While intPos > 0
            If Mid$(strFilename, intPos, 1) <> gstrSEP_DIR Then
                intPos = intPos - 1
            Else
                Exit Do
            End If
        Loop
    End If

    If intPos > 0 Then
        strPathOnly = Left$(strFilename, intPos)
        If Right$(strPathOnly, 1) = gstrCOLON Then
            strPathOnly = strPathOnly & gstrSEP_DIR
        End If
    Else
        strPathOnly = CurDir$
    End If

    If Right$(strPathOnly, 1) = gstrSEP_DIR Then
        strPathOnly = Left$(strPathOnly, Len(strPathOnly) - 1)
    End If

    GetPathName = UCase16(strPathOnly)
    
    Err = 0
End Function

'Devuelve la ruta de acceso para la raíz del primer disco fijo
Function strRootDrive() As String
    Dim intDriveNum As Integer
    
    For intDriveNum = 0 To Asc("Z") - Asc("A") - 1
        If GetDriveType(intDriveNum) = intDRIVE_FIXED Then
            strRootDrive = Chr$(Asc("A") + intDriveNum) & gstrCOLON & gstrSEP_DIR
            Exit Function
        End If
    Next intDriveNum
    
    strRootDrive = "C:\"
End Function

'Devuelve "" si la ruta de acceso no está completa o si es un nombre UNC
Function strGetDriveFromPath(ByVal strPath As String) As String
    If Len(strPath) < 2 Then
        Exit Function
    End If
    
    If Mid$(strPath, 2, 1) <> gstrCOLON Then
        Exit Function
    End If
    
    strGetDriveFromPath = Mid$(strPath, 1, 1) & gstrCOLON & gstrSEP_DIR
End Function

Public Function fValidFilename(strFilename As String) As Boolean
'
' Esta rutina comprueba que strFileName es un nombre de archivo válido.
' Comprueba que su longitud es menor que el máximo permitido y que
' no contiene ningún carácter no válido.
'
    If Not fCheckFNLength(strFilename) Then
        '
        ' El nombre es demasiado largo.
        '
        fValidFilename = False
        Exit Function
    End If
    '
    ' Busca en la lista de caracteres no válidos para nombres de archivo y
    ' se asegura de ninguno de ellos está incluido en la cadena.
    '
    Dim iInvalidChar As Integer
    Dim iFilename As Integer
    Dim strInvalidChars As String
    
    strInvalidChars = ResolveResString(resCOMMON_INVALIDFILECHARS)
    
    For iInvalidChar = 1 To Len(strInvalidChars)
        If InStr(strFilename, Mid$(strInvalidChars, iInvalidChar, 1)) <> 0 Then
            fValidFilename = False
            Exit Function
        End If
    Next iInvalidChar
    
    fValidFilename = True
    
End Function
Public Function fValidNTGroupName(strGroupName) As Boolean
'
' Esta rutina comprueba que strGroupName es un nombre válido de grupo.
' Comprueba que su longitud es menor que el máximo permitido y que
' no contiene ningún carácter no válido.
'
    If Len(strGroupName) > gintMAX_GROUPNAME_LEN Then
        fValidNTGroupName = False
        Exit Function
    End If
    '
    ' Busca en la lista de caracteres no válidos para nombres de archivo y
    ' se asegura de ninguno de ellos está incluido en la cadena.
    '
    Dim iInvalidChar As Integer
    Dim iFilename As Integer
    Dim strInvalidChars As String
    
    strInvalidChars = ResolveResString(resGROUPINVALIDCHARS)
    
    For iInvalidChar = 1 To Len(strInvalidChars)
        If InStr(strGroupName, Mid$(strInvalidChars, iInvalidChar, 1)) <> 0 Then
            fValidNTGroupName = False
            Exit Function
        End If
    Next iInvalidChar
    
    fValidNTGroupName = True
    
End Function
'-----------------------------------------------------------
' Procedimiento: CountIcons
'
' Determina cuántos iconos se deben instalar, contándolos
' en el archivo de información de la instalación (SETUP.LST)
'-----------------------------------------------------------
'
Function CountIcons(ByVal strSection As String) As Integer
    Dim intIdx As Integer
    Dim sFile As FILEINFO
    Dim cIcons As Integer

    '
    'Para cada archivo de la sección especificada, lee información del
    'archivo de información de la instalación
    '
    intIdx = 1
    cIcons = 0
    Do While ReadSetupFileLine(strSection, intIdx, sFile) = True
        If Len(strUnQuoteString(sFile.strProgramIconTitle)) > 0 Then
            cIcons = cIcons + 1
        End If
        intIdx = intIdx + 1
    Loop
    CountIcons = cIcons
End Function
'-----------------------------------------------------------
' Procedimiento: CreateIcons
'
' Recorre la lista de archivos de SETUP.LST y crea iconos en
' el grupo de programas para los archivos que los necesitan.
'-----------------------------------------------------------
'
Sub CreateIcons(ByVal strSection As String, ByVal strGroupName As String)
    Dim intIdx As Integer
    Dim sFile As FILEINFO
    Dim strProgramIconTitle As String
    Dim strProgramIconCmdLine As String
    Dim strProgramPath As String
    Dim strProgramArgs As String
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Const CompareBinary = 0
    '
    'Para cada archivo de la sección especificada, lee información del
    'archivo de información de la instalación
    '
    intIdx = 1
    Do While ReadSetupFileLine(strSection, intIdx, sFile) = True
        '
        ' Obtiene el título y la línea de comandos del icono
        '
        strProgramIconTitle = sFile.strProgramIconTitle
        strProgramIconCmdLine = sFile.strProgramIconCmdLine
        '
        ' si se especificó ProgramIcon, creamos un icono;
        ' de lo contrario, no lo creamos.
        '
        If Trim(strUnQuoteString(strProgramIconTitle)) <> "" Then
            '
            ' Si la línea de comandos no se especifica en SETUP.LST y el icono sí,
            ' usamos la ruta de destino de los archivos como línea de comandos.
            ' En este caso no hay parámetros.
            '
            If Trim(strUnQuoteString(strProgramIconCmdLine)) = "" Then
                strProgramPath = sFile.strDestDir & gstrSEP_DIR & sFile.strDestName
                strProgramArgs = ""
            Else
                '
                ' Analiza la línea de comandos para determinar cuál es el .exe, etc.
                ' y cuáles son los parámetros. El primer espacio en blanco que no
                ' está entre comillas marca el final del .exe, etc. Todo lo que haya
                ' detrás son parámetros o argumentos para el .exe. NOTA: es importante
                ' que si el .exe está entre comillas los parámetros no estén dentro
                ' de las mismas comillas. Cada argumento puede estar entre comillas
                ' siempre y cuando no estén en las mimas comillas que el .exe.
                '
                intAnchor = 1
                intOffset = intGetNextFldOffset(intAnchor, strProgramIconCmdLine, " ", CompareBinary)
                If intOffset = 0 Then intOffset = Len(strProgramIconCmdLine) + 1
                strProgramPath = Trim(Left(strProgramIconCmdLine, intOffset - 1))
                '
                ' Obtuvo el .exe; ahora obtiene los parámetros.
                '
                strProgramArgs = Trim(Mid(strProgramIconCmdLine, intOffset + 1))
            End If
            '
            ' Expande todas las macros de directorio de destino incluidas en
            ' la ruta de acceso del programa y los argumentos
            '
            strProgramPath = ResolveDestDir(strProgramPath)
            strProgramArgs = ResolveDestDirs(strProgramArgs)
            '
            ' Ya tenemos todo lo necesario; creamos el icono.
            '
            CreateOSLink frmSetup1, strGroupName, strProgramPath, strProgramArgs, strProgramIconTitle
        ElseIf Trim(strUnQuoteString(strProgramIconCmdLine)) <> "" Then
            '
            ' Este archivo contenía una línea de comandos de SETUP.LST pero ningún icono.
            ' Es un error. Informamos al usuario y saltamos este icono o anulamos.

            '
            If gfNoUserInput Or MsgWarning(ResolveResString(resICONMISSING, "|1", sFile.strDestName), vbYesNo Or vbExclamation, gstrSETMSG) = vbNo Then
                ExitSetup frmSetup1, gintRET_FATAL
            End If
        End If
        intIdx = intIdx + 1
    Loop
End Sub

