Attribute VB_Name = "basCommon"
Option Explicit
Option Compare Text

'
' Constantes globales
'
Global Const gstrNULL$ = ""                             ' Cadena vacía

Global Const gstrSEP_DIR$ = "\"                         ' Carácter separador de directorios
Public Const gstrSEP_REGKEY$ = "\"                      ' Carácter de separación de claves del Registro.
Global Const gstrSEP_DRIVE$ = ":"                       ' Carácter de separación de unidad, p.e., C:\
Global Const gstrSEP_DIRALT$ = "/"                      ' Carácter alternativo de separación de directorio
Global Const gstrSEP_EXT$ = "."                         ' Carácter de separación de extensión de archivo
Public Const gstrSEP_PROGID = "."
Public Const gstrSEP_FILE$ = "|"                        ' Usa el carácter para delimitar listas de nombres de archivo ya que no es un carácter válido en un nombre de archivo.
Public Const gstrSEP_LIST = "|"
Public Const gstrSEP_URL$ = "://"                       ' Separador que sigue a HPPT en una dirección URL
Public Const gstrSEP_URLDIR$ = "/"                      ' Separador para separar directorios en una dirección URL.

Global Const gstrUNC$ = "\\"                            ' Especificador UNC \\
Global Const gstrCOLON$ = ":"
Global Const gstrSwitchPrefix1 = "-"
Global Const gstrSwitchPrefix2 = "/"
Global Const gstrCOMMA$ = ","
Global Const gstrDECIMAL$ = "."
Global Const gstrQUOTE$ = """"
Public Const gstrCCOMMENT$ = "//"                       ' Indicador de comentario usado en C, etc.
Public Const gstrASSIGN$ = "="
Global Const gstrINI_PROTOCOL = "Protocol"
Public Const gstrREMOTEAUTO = "RA"
Public Const gstrDCOM = "DCOM"

Global Const gintMAX_SIZE% = 255                        ' Máximo tamaño del búfer
Global Const gintMAX_PATH_LEN% = 260                    ' Longitud de ruta máxima permitida para nombre de archivo
                                                        ' y argumentos de la línea de comandos para NT (Intel) y Win95.
Global Const gintMAX_GROUPNAME_LEN% = 30                ' Máxima longitud permitida para un nombre de grupo de NT 3.51.
Global Const gintMIN_BUTTONWIDTH% = 1200
Global Const gsngBUTTON_BORDER! = 1.4

Global Const intDRIVE_REMOVABLE% = 2                    'Constantes de GetDriveType
Global Const intDRIVE_FIXED% = 3
Global Const intDRIVE_REMOTE% = 4
Global Const intDRIVE_CDROM% = 5
Global Const intDRIVE_RAMDISK% = 6

Global Const gintNOVERINFO% = 32767                     'indicador del número de la versión

'Nombre de archivos
Global Const gstrFILE_SETUP$ = "SETUP.LST"              'Nombre del archivo de información de instalación
Public Const gstrTEMP_DIR$ = "TEMP"
Public Const gstrTMP_DIR$ = "TMP"

'Tipos de macros compartidas para archivos
Global Const mstrPRIVATEFILE = ""
Global Const mstrSHAREDFILE = "$(Shared)"

'Claves de archivos INI
Global Const gstrINI_SETUP$ = "Setup"
Global Const gstrINI_APPNAME$ = "Title"
Global Const gstrINI_APPDIR$ = "DefaultDir"
Global Const gstrINI_APPEXE$ = "AppExe"
Public Const gstrINI_APPTOUNINSTALL = "AppToUninstall"
Global Const gstrINI_APPPATH$ = "AppPath"
Global Const gstrINI_FORCEUSEDEFDEST = "ForceUseDefDir"
Global Const gstrINI_DEFGROUP$ = "DefProgramGroup"

Public Const gstrEXT_DEP$ = "DEP"

'Macros de archivos de información de la instalación
Global Const gstrAPPDEST$ = "$(AppPath)"
Global Const gstrWINDEST$ = "$(WinPath)"
Global Const gstrWINSYSDEST$ = "$(WinSysPath)"
Global Const gstrWINSYSDESTSYSFILE$ = "$(WinSysPathSysFile)"
Global Const gstrPROGRAMFILES$ = "$(ProgramFiles)"
Global Const gstrCOMMONFILES$ = "$(CommonFiles)"
Global Const gstrCOMMONFILESSYS$ = "$(CommonFilesSys)"
Global Const gstrDAODEST$ = "$(MSDAOPath)"
Public Const gstrDONOTINSTALL$ = "$(DoNotInstall)"

'Contantes de puntero del mouse
Global Const gintMOUSE_DEFAULT% = 0
Global Const gintMOUSE_HOURGLASS% = 11

'Constantes de MsgError()
Global Const MSGERR_ERROR = 1
Global Const MSGERR_WARNING = 2

'Constantes de MsgBox
Global Const MB_OK = 0                                  'Sólo el botón Aceptar
Global Const MB_OKCANCEL = 1                            'Botones Aceptar y Cancelar
Global Const MB_ABORTRETRYIGNORE = 2                    'Botones Anular, Reintentar e Ignorar
Global Const MB_YESNO = 4                               'Botones Sí y No
Global Const MB_RETRYCANCEL = 5                         'Botones Reintentar y Cancelar
Global Const MB_ICONSTOP = 16                           'Mensaje crítico
Global Const MB_ICONQUESTION = 32                       'Consulta de advertencia
Global Const MB_ICONEXCLAMATION = 48                    'Mensaje de advertencia
Global Const MB_ICONINFORMATION = 64                    'Mensaje de información
Global Const MB_DEFBUTTON1 = 0                          'Primer botón predeterminado
Global Const MB_DEFBUTTON2 = 256                        'Segundo botón predeterminado
Global Const MB_DEFBUTTON3 = 512                        'Tercer botón predeterminado

'Valores de retorno de MsgBox
Global Const IDOK = 1                                   'Botón Aceptar presionado
Global Const IDCANCEL = 2                               'Botón Cancelar presionado
Global Const IDABORT = 3                                'Botón Abortar presionado
Global Const IDRETRY = 4                                'Botón Reintentar presionado
Global Const IDIGNORE = 5                               'Botón Ignorar presionado
Global Const IDYES = 6                                  'Botón Sí presionado
Global Const IDNO = 7                                   'Botón No presionado

'
'Definiciones de tipo
'
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    nReserved1 As Integer
    nReserved2 As Integer
    szPathName As String * 256
End Type

Type VERINFO                                            'FIXEDFILEINFO de versión
    strPad1 As Long                                     'Estructura para la versión
    strPad2 As Long                                     'Estructura para la firma
    nMSLo As Integer                                    'Palabra inferior de la versión en # MS DWord
    nMSHi As Integer                                    'Palabra superior de la versión en # MS DWord
    nLSLo As Integer                                    'Palabra inferior de la versión en # LS DWord
    nLSHi As Integer                                    'Palabra superior de la versión en # LS DWord
    strPad3(1 To 16) As Byte                            'Salta 16 bytes en la estructura de VERINFO
    FileOS As Long                                      'Información sobre el SO de destino para este archivo.
    strPad4(1 To 16) As Byte                            'Estructura del resto de VERINFO (16 bytes)
End Type

Type PROTOCOL
    strName As String
    strFriendlyName As String
End Type

Type OSVERSIONINFO 'para la llamada a la función GetVersionEx de la API
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Global Const OF_EXIST& = &H4000&
Global Const OF_SEARCH& = &H400&
Global Const HFILE_ERROR% = -1

'
'Variables globales
'
Global LF$                                              'salto con una línea
Global LS$                                              'salto con dos líneas
Public CRLF As String                                   'retorno de carro y avance de línea
'
' Variables globales usadas para instalaciones silenciosas y SMS
'
Public gfSilent As Boolean                              ' Indica si se realiza una instalación silenciosa
Public gstrSilentLog As String                          ' nombre de archivo de salida para instalación silenciosa.
Public gfSMS As Boolean                                 ' Indica si se realiza una instalación SMS
Public gstrMIFFile As String                            ' estado del archivo de salida para SMS
Public gfSMSStatus As Boolean                           ' estado de la instalación para SMS
Public gstrSMSDescription As String                     ' cadena de descripción que se escribe en un archivo MIF durante una instalación SMS
Public gfNoUserInput As Boolean                         ' True si gfSMS o gfSilent son True
Public gfDontLogSMS As Boolean                          ' Impide que se cargue MsgFunc en SMS (p.e., para mensajes de confirmación)
Public Const MAX_SMS_DESCRIP = 255                      ' SMS no es compatible con cadenas de descripción mayores de 255 caracteres.
'
'Lista de protocolos disponibles
'
Global gProtocol() As PROTOCOL
Global gcProtocols As Integer
'
' Son necesarios AXDist.exe y wint351.exe. Son EXE de extracción automática
' que instalan otros archivos que no instala setup1.
'
Public gfAXDist As Boolean
Public Const gstrFILE_AXDIST = "AXDIST.EXE"
Public gstrAXDISTInstallPath As String
Public gfAXDistChecked As Boolean
Public gfWINt351 As Boolean
Public Const gstrFILE_WINT351 = "WINt351.EXE"
Public gstrWINt351InstallPath As String
Public gfWINt351Checked As Boolean
'
'Declaraciones DLL de API del SetupToolkit de 32 bits
'
Declare Function DiskSpaceFree Lib "VB5STKIT.DLL" Alias "DISKSPACEFREE" () As Long
Declare Function SetTime Lib "VB5STKIT.DLL" (ByVal strFileGetTime As String, ByVal strFileSetTime As String) As Integer
Declare Function AllocUnit Lib "VB5STKIT.DLL" () As Long
Declare Function GetWinPlatform Lib "VB5STKIT.DLL" () As Long
Declare Function fNTWithShell Lib "VB5STKIT.DLL" () As Boolean
Declare Function FSyncShell Lib "VB5STKIT.DLL" Alias "SyncShell" (ByVal strCmdLine As String, ByVal intCmdShow As Long) As Long
Declare Function DLLSelfRegister Lib "VB5STKIT.DLL" (ByVal lpDllName As String) As Integer
Declare Function GetClsidFromActXFile Lib "VB5STKIT.DLL" (ByVal pszFilename As String, ByVal pszProgID As String, ByVal pszClsid As String) As Long
Declare Function RegisterTLB Lib "VB5STKIT.DLL" (ByVal lpTLBName As String) As Integer
Declare Sub lmemcpy Lib "VB5STKIT.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
Declare Function OSfCreateShellGroup Lib "VB5STKIT.DLL" Alias "fCreateShellFolder" (ByVal lpstrDirName As String) As Long
Declare Function OSfCreateShellLink Lib "VB5STKIT.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function OSfRemoveShellLink Lib "VB5STKIT.DLL" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
Private Declare Function OSGetLongPathName Lib "VB5STKIT.DLL" Alias "GetLongPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare Function OpenFile Lib "kernel32" (ByVal lpFilename As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
Declare Function GetTempFilename32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFilename As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_ERR = (-1)

Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, lVerHandle As Long) As Long
Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'-----------------------------------------------------------
' Procedimiento: AddDirSep
' Agrega un separador de ruta de directorio (barra invertida) al
' final de la ruta si no existe
'
' Entrada/Salida: [strPathName] - ruta para agregar el separador
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
'-----------------------------------------------------------
' Procedimiento: AddURLDirSep
' Agrega un separador de ruta URL (barra normal) al final
' de una dirección URL si no existe
'
' Entrada/Salida: [strPathName] - ruta a agregar el separador
'-----------------------------------------------------------
'
Sub AddURLDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = Trim(strPathName) & gstrSEP_URLDIR
    End If
End Sub

'-----------------------------------------------------------
' Función: FileExists
' Determina si existe el archivo especificado
'
' Entrada: [strPathName] - archivo a comprobar
'
' Devuelve: True si existe, si no False
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    ' Si la cadena tiene comillas las elimina.
    '
    strPathName = strUnQuoteString(strPathName)
    '
    'Elimina cualquier carácter separador de directorio
    '
    If Right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Intenta abrir el archivo, el valor de retorno de esta función es False
    'en caso contrario es True
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function

'-----------------------------------------------------------
' Función: DirExists
'
' Determina si el nombre de directorio especificado existe.
' Esta función se puede usar para determinar si un disco de
' instalación está en la unidad pasando
' algo como 'A:\'.
'
' Entrada: [strDirName] - nombre del directorio a comprobar
'
' Devuelve: True si el directorio existe, False en caso contrario
'-----------------------------------------------------------
'
Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    AddDirSep strDirName
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = gstrNULL)

    Err = 0
End Function

'-----------------------------------------------------------
' Función: GetDriveType
' Determina si un disco es fijo, removible, etc. llamando
' a GetDriveType() de Windows
'-----------------------------------------------------------
'
Function GetDriveType(ByVal intDriveNum As Integer) As Integer
    '
    ' Esta función espera un número de unidad entero en Win16 o una cadena en Win32
    '
    Dim strDriveName As String
    
    strDriveName = Chr$(Asc("A") + intDriveNum) & gstrSEP_DRIVE & gstrSEP_DIR
    GetDriveType = CInt(GetDriveType32(strDriveName))
End Function

'-----------------------------------------------------------
' Función: ReadProtocols
' Lee los protocolos disponibles del archivo especificado.
'
' Entrada: [strInputFilename] - Nombre de archivo INI del que leer protocolos
'          [strINISection] - Nombre de la sección INI
'-----------------------------------------------------------
Function ReadProtocols(ByVal strInputFilename As String, ByVal strINISection As String) As Boolean
    Dim intIdx As Integer
    Dim fOk As Boolean
    Dim strInfo As String
    Dim intOffset As Integer
    
    intIdx = 0
    fOk = True
    Erase gProtocol
    gcProtocols = 0
    
    Do
        strInfo = ReadIniFile(strInputFilename, strINISection, gstrINI_PROTOCOL & Format$(intIdx + 1))
        If strInfo <> gstrNULL Then
            intOffset = InStr(strInfo, gstrCOMMA)
            If intOffset > 0 Then
                'El nombre "incorrecto" será lo primero de la línea
                ReDim Preserve gProtocol(intIdx + 1)
                gcProtocols = intIdx + 1
                gProtocol(intIdx + 1).strName = Left$(strInfo, intOffset - 1)
                
                '... seguido por el nombre amistoso
                gProtocol(intIdx + 1).strFriendlyName = Mid$(strInfo, intOffset + 1)
                If (gProtocol(intIdx + 1).strName = "") Or (gProtocol(intIdx + 1).strFriendlyName = "") Then
                    fOk = False
                End If
            Else
                fOk = False
            End If

            If Not fOk Then
                Exit Do
            Else
                intIdx = intIdx + 1
            End If
        End If
    Loop While strInfo <> gstrNULL
    
    ReadProtocols = fOk
End Function

'-----------------------------------------------------------
' Función: ResolveResString
' Lee recursos y reemplaza las macros por valores
'
' Ejemplo, con el número de recurso 14:
'    "Imposible leer '|1' en la unidad |2"
'   la llamada
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   Devolvería la cadena
'     "Imposible leer 'TXTFILE.TXT' en la unidad A:"
'
' Entrada: [resID] - identificador de recurso
'          [varReplacements] - parejas de macro y valor
'-----------------------------------------------------------
'
Public Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    
    strResString = LoadResString(resID)
    
    ' Para cada pareja de valor y de macro que se pasa...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Reemplaza las apariciones de strMacro con strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function
'-----------------------------------------------------------
' Procedimiento: GetLicInfoFromVBL
' Procesa un nombre de archivo VBL y extrae la clave de licencia
' del Registro y la información de licencia.
'
' Entrada: [strVBLFile] - debe ser un VBL válido.
'
' Salida: [strLicKey] - clave del registro dónde escribir la información de licencia.
'                       Esta clave se agregará a
'                       HKEY_CLASSES_ROOT\Licenses. Es un
'                       guid.
' Salida: [strLicVal] - información de licencia. Normalmente en forma de cadena con
'                       caracteres codificados.
'-----------------------------------------------------------
'
Public Sub GetLicInfoFromVBL(strVBLFile As String, strLicKey As String, strLicVal As String)
    Dim fn As Integer
    Const strREGEDIT = "REGEDIT"
    Const strLICKEYBASE = "HKEY_CLASSES_ROOT\Licenses\"
    Dim strTemp As String
    Dim posEqual As Integer
    Dim fLicFound As Boolean
    
    fn = FreeFile
    Open strVBLFile For Input Access Read Lock Read Write As #fn
    '
    ' Lee el archivo hasta encontrar un línea que comience con strLICKEYBASE
    '
    fLicFound = False
    Do While Not EOF(fn)
        Line Input #fn, strTemp
        strTemp = Trim(strTemp)
        If Left$(strTemp, Len(strLICKEYBASE)) = strLICKEYBASE Then
            '
            ' Tenemos la línea que queremos.
            '
            fLicFound = True
            Exit Do
        End If
    Loop

    Close fn
    
    If fLicFound Then
        '
        ' Procesa el dato de esta línea para separar
        ' la clave y la información de licencia. La línea
        ' debe tener el formato:
        ' "HKEY_CLASSES_ROOT\Licenses\<lickey> = <licval>"
        '
        posEqual = InStr(strTemp, gstrASSIGN)
        If posEqual > 0 Then
            strLicKey = Mid$(Trim(Left$(strTemp, posEqual - 1)), Len(strLICKEYBASE) + 1)
            strLicVal = Trim(Mid$(strTemp, posEqual + 1))
        End If
    Else
        strLicKey = gstrNULL
        strLicVal = gstrNULL
    End If
End Sub

 '-----------------------------------------------------------
 ' Función: GetLongPathName
 '
 ' Recupera el nombre de ruta largo de una ruta que contiene
 ' nombres de archivos o subdirectorios en formato corto
 '-----------------------------------------------------------
 '
 Function GetLongPathName(ByVal strShortPath As String) As String
    Const cchBuffer = 300
    Dim strLongPath As String
    Dim lResult As Long
    
    On Error GoTo 0
    
    strLongPath = String(cchBuffer, Chr$(0))
    lResult = OSGetLongPathName(strShortPath, strLongPath, cchBuffer)
    If lResult = 0 Then
        Error 53 ' Archivo no encontrado
    Else
        GetLongPathName = StripTerminator(strLongPath)
    End If
 End Function
 
 '-----------------------------------------------------------
 ' Función: GetShortPathName
 '
 ' Recupera el nombre de ruta corto de una ruta que contiene
 ' nombres de archivos o subdirectorios en formato largo
 '-----------------------------------------------------------
 '
 Function GetShortPathName(ByVal strLongPath As String) As String
     Const cchBuffer = 300
     Dim strShortPath As String
     Dim lResult As Long

     On Error GoTo 0
     strShortPath = String(cchBuffer, Chr$(0))
     lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
     If lResult = 0 Then
         Error 53 ' Archivo no encontrado
     Else
         GetShortPathName = StripTerminator(strShortPath)
     End If
 End Function
 
'-----------------------------------------------------------
' Función: GetTempFilename
' Obtiene un nombre de archivo temporal para una unidad especificada
' y un prefijo de nombre de archivo
' Parámetros:
'   strDestPath - Ubicación dónde se creará el archivo temporal. Si es
'                 una cadena vacía, se usará la ubicación especificada en
'                 las variables de entorno tmp o temp.
'   lpPrefixString - Los primeros tres caracteres de esta cadena formarán
'                    parte del nombre del archivo temporal que se devuelve.
'   wUnique - Se establece a 0 para crear un nombre de archivo único. Se puede,
'             establecer con un Integer, en este caso el nombre del archivo temporal
'             se devolverá con este Integer formando parte del nombre.
'   lpTempFilename - El nombre del archivo temporal se devuelve en esta variable.
' Devuelve:
'   True si tiene éxito la función; False en caso contrario
'-----------------------------------------------------------
'
Function GetTempFilename(ByVal strDestPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFilename As String) As Boolean
    If strDestPath = gstrNULL Then
        '
        ' No se especificó un destino, se usa el directorio temporal.
        '
        strDestPath = String(gintMAX_PATH_LEN, vbNullChar)
        If GetTempPath(gintMAX_PATH_LEN, strDestPath) = 0 Then
            GetTempFilename = False
            Exit Function
        End If
    End If
    lpTempFilename = String(gintMAX_PATH_LEN, vbNullChar)
    GetTempFilename = GetTempFilename32(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
'-----------------------------------------------------------
' Función: GetDefMsgBoxButton
' Decodifica los indicadores que se pasan a la función MsgBox para
' determinar qué botón es el predeterminado. Se usa en instalaciones
' silenciosas.
'
' Entrada: [intFlags] - Indicadores de MsgBox
'
' Devuelve: Número de botones definidos en
'               vbOK        1   Botón Aceptar presionado.
'               vbCancel    2   Botón Cancelar presionado.
'               vbAbort     3   Botón Anular presionado.
'               vbRetry     4   Botón Reintentar presionado.
'               vbIgnore    5   Botón Ignorar presionado.
'               vbYes       6   Botón Sí presionado.
'               vbNo        7   Botón No presionado.
'-----------------------------------------------------------
'
Function GetDefMsgBoxButton(intFlags) As Integer
    '
    ' Primero se determina el ordinal del botón predeterminado
    ' en el cuadro de mensaje.
    '
    Dim intButtonNum As Integer
    Dim intDefButton As Integer
    
    If (intFlags And vbDefaultButton2) = vbDefaultButton2 Then
        intButtonNum = 2
    ElseIf (intFlags And vbDefaultButton3) = vbDefaultButton3 Then
        intButtonNum = 3
    Else
        intButtonNum = 1
    End If
    '
    ' Ahora se determina el tipo de cuadro de mensaje con el que estamos
    ' tratando y se devuelve el botón predeterminado.
    '
    If (intFlags And vbRetryCancel) = vbRetryCancel Then
        intDefButton = IIf(intButtonNum = 1, vbRetry, vbCancel)
    ElseIf (intFlags And vbYesNoCancel) = vbYesNoCancel Then
        Select Case intButtonNum
            Case 1
                intDefButton = vbYes
            Case 2
                intDefButton = vbNo
            Case 3
                intDefButton = vbCancel
            'End Case
        End Select
    ElseIf (intFlags And vbOKCancel) = vbOKCancel Then
        intDefButton = IIf(intButtonNum = 1, vbOK, vbCancel)
    ElseIf (intFlags And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
        Select Case intButtonNum
            Case 1
                intDefButton = vbAbort
            Case 2
                intDefButton = vbRetry
            Case 3
                intDefButton = vbIgnore
            'End Case
        End Select
    ElseIf (intFlags And vbYesNo) = vbYesNo Then
        intDefButton = IIf(intButtonNum = 1, vbYes, vbNo)
    Else
        intDefButton = vbOK
    End If
    
    GetDefMsgBoxButton = intDefButton
    
End Function
'-----------------------------------------------------------
' Función: GetDiskSpaceFree
' Obtiene la cantidad de espacio en disco libre de la unidad especificada
'
' Entrada: [strDrive] - unidad en la que comprobar el espacio
'
' Devuelve: Cantidad de espacio libre en disco, o -1 si ha sucedido un error
'-----------------------------------------------------------
'
Function GetDiskSpaceFree(ByVal strDrive As String) As Long
    Dim strCurDrive As String
    Dim lDiskFree As Long

    On Error Resume Next

    '
    'Guarda la unidad actual
    '
    strCurDrive = Left$(CurDir$, 2)

    '
    'Ajusta la unidad para que sólo tenga una letra y una barra
    '
    If InStr(strDrive, gstrSEP_DRIVE) = 0 Or Len(strDrive) > 2 Then
        strDrive = Left$(strDrive, 1) & gstrSEP_DRIVE
    End If

    '
    'Cambia a la unidad de la que queremos saber el espacio. La función DiskSpaceFree()
    'de la API sólo funciona en la unidad actual.
    '
    ChDrive strDrive

    '
    'Si es imposible cambiar a la unidad solicitada, es un error, de lo contrario
    'se devuelve la cantidad en disco disponible
    '
    If Err <> 0 Or (strDrive <> Left$(CurDir$, 2)) Then
        lDiskFree = -1
    Else
        lDiskFree = DiskSpaceFree()
        If Err <> 0 Then    'Si es imposible encontrar la DLL del instalador
            lDiskFree = -1
        End If
    End If

    If lDiskFree = -1 Then
        MsgError Error$ & LS$ & ResolveResString(resDISKSPCERR) & strDrive, MB_ICONEXCLAMATION, gstrTitle
    End If

    GetDiskSpaceFree = lDiskFree

    '
    'Se restablece estableciendo la unidad con la original
    '
    ChDrive strCurDrive

    Err = 0
End Function

'-----------------------------------------------------------
' Función: GetUNCShareName
'
' Con un nombre UNC, devuelve la parte más a la izquierda
' del directorio que representa el nombre del equipo y el nombre compartido.
' P.e., con "\\SCHWEIZ\PUBLIC\APPS\LISTING.TXT", devuelve
' la cadena "\\SCHWEIZ\PUBLIC"
'
' Devuelve una cadena que representa el nombre del equipo si
'   la ruta tiene un nombre válido, si no devuelve NULL
'-----------------------------------------------------------
'
Function GetUNCShareName(ByVal strFN As String) As Variant
    GetUNCShareName = Null
    If IsUNCName(strFN) Then
        Dim iFirstSeparator As Integer
        iFirstSeparator = InStr(3, strFN, gstrSEP_DIR)
        If iFirstSeparator > 0 Then
            Dim iSecondSeparator As Integer
            iSecondSeparator = InStr(iFirstSeparator + 1, strFN, gstrSEP_DIR)
            If iSecondSeparator > 0 Then
                GetUNCShareName = Left$(strFN, iSecondSeparator - 1)
            Else
                GetUNCShareName = strFN
            End If
        End If
    End If
End Function

'-----------------------------------------------------------
' Función: GetWindowsSysDir
'
' Llama a la API de Windows para obtener el directorio windows\SYSTEM
' y comprueba que tiene un símbolo de separación
'
' Devuelve: El directorio windows\SYSTEM
'-----------------------------------------------------------
'
Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Obtiene el directorio del sistema y recorta el búfer para que la longitud coincida
    'y agrega un separador de directorio (barra invertida) si el API no lo devuelve
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        GetWindowsSysDir = strBuf
    Else
        GetWindowsSysDir = gstrNULL
    End If
End Function
'-----------------------------------------------------------
' Procedimiento: TreatAsWin95
'
' Devuelve True si estamos ejecutando Windows 95
' o si estamos tratando esta versión de NT como si fuese
' Windows 95 para el registro y el inicio de la aplicación
' y cuando se desee desinstalar.
'-----------------------------------------------------------
'
Function TreatAsWin95() As Boolean
    If IsWindows95() Then
        TreatAsWin95 = True
    ElseIf fNTWithShell() Then
        TreatAsWin95 = True
    Else
        TreatAsWin95 = False
    End If
End Function
'-----------------------------------------------------------
' Función: IsDepFile
'
' Devuelve True si el archivo que se pasa a está función es un
' archivo de dependencias (*.dep). Hacemos esta comprobación
' comprobando que extensión es .dep y que contiene la
' información de la versión.
'-----------------------------------------------------------
'
Function fIsDepFile(strFilename As String) As Boolean
    Const strEXT_DEP = "DEP"
    
    fIsDepFile = False
    
    If UCase(Extension(strFilename)) = strEXT_DEP Then
        If GetFileVersion(strFilename) <> gstrNULL Then
            fIsDepFile = True
        End If
    End If
End Function

'-----------------------------------------------------------
' Función: IsWin32
'
' Devuelve True si el programa se está ejecutando en Win32 (o
'   en un sistema operativo de 32 bits)
'-----------------------------------------------------------
'
Function IsWin32() As Boolean
    IsWin32 = (IsWindows95() Or IsWindowsNT())
End Function

'-----------------------------------------------------------
' Función: IsWindows95
'
' Devuelve True si este programa se está ejecutando en Windows 95
'   o posterior
'-----------------------------------------------------------
'
Function IsWindows95() As Boolean
    Const dwMask95 = &H2&
    If GetWinPlatform() And dwMask95 Then
        IsWindows95 = True
    Else
        IsWindows95 = False
    End If
End Function

'-----------------------------------------------------------
' Función: IsWindowsNT
'
' Devuelve True si este programa se está ejecutando en Windows NT
'-----------------------------------------------------------
'
Function IsWindowsNT() As Boolean
    Const dwMaskNT = &H1&
    If GetWinPlatform() And dwMaskNT Then
        IsWindowsNT = True
    Else
        IsWindowsNT = False
    End If
End Function

'-----------------------------------------------------------
' Función: IsWindowsNT4WithoutSP2
'
' Determina si el usuario está ejecutando Windows NT 4.0
' sin el Service Pack 2 (SP2). Si lo está ejecutando en cualquier
' otra plataforma, devuelve False.
'
' Entrada: [ninguno]
'
' Devuelve: True si está ejecutando Windows NT 4.0
' sin el Service Pack 2 instalado.
'-----------------------------------------------------------
'
Function IsWindowsNT4WithoutSP2() As Boolean
    IsWindowsNT4WithoutSP2 = False
    
    If Not IsWindowsNT() Then
        Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    '¿Es Windows NT 4.0?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion <> NT4MajorVersion) Or (osvi.dwMinorVersion <> NT4MinorVersion) Then
        'No.  Devuelve False.
        Exit Function
    End If
    
    'Si no está instalado un service pack, o si está instalado el Service Pack 1
    'devuelve True.
    Const strSP1 = "SERVICE PACK 1"
    If strCSDVersion = "" Then
        IsWindowsNT4WithoutSP2 = True 'No hay un service pack instalado
    ElseIf strCSDVersion = strSP1 Then
        IsWindowsNT4WithoutSP2 = True 'Sólo está instalado el Service Pack 1
    End If
End Function

'-----------------------------------------------------------
' Función: IsUNCName
'
' Determina si el nombre de ruta especificado es un nombre UNC.
' Los nombres UNC (Universal Naming Convention) se usan normalmente
' para especificar recursos de equipos, como unidades de red
' compartidas, canalizaciones con nombre, etc. Un ejemplo de nombres UNC es
' "\\SERVIDOR\COMPARTIDO\ARCHIVO.EXT".
'
' Entrada: [strPathName] - ruta para comprobar
'
' Devuelve: True si la ruta es un nombre UNC, False en caso contrario
'-----------------------------------------------------------
'
Function IsUNCName(ByVal strPathName As String) As Integer
    Const strUNCNAME$ = "\\//\"        'para comprobar \\, //, \/, /\

    IsUNCName = ((InStr(strUNCNAME, Left$(strPathName, 2)) > 0) And _
                 (Len(strPathName) > 1))
End Function
'-----------------------------------------------------------
' Función: LogSilentMsg
'
' Si es una instalación silenciosa, esta rutina escribe un mensaje en
' en archivo gstrSilentLog.
'
' Entrada: [strMsg] - El mensaje
'
' Normalmente, se llama a esta rutina en lugar de visualizar un
' MsgBox, strMsg es el mismo mensaje que debería aparecer
' en el MsgBox

'-----------------------------------------------------------
'
Sub LogSilentMsg(strMsg As String)
    If Not gfSilent Then Exit Sub
    
    Dim fn As Integer
    
    On Error Resume Next
    
    fn = FreeFile
    
    Open gstrSilentLog For Append As fn
    Print #fn, strMsg
    Close fn
    Exit Sub
End Sub
'-----------------------------------------------------------
' Función: LogSMSMsg
'
' Si es una instalación SMS, esta rutina agrega un mensaje
' a la cadena gstrSMSDescription. Esta cadena se escribirá
' posteriormente en el archivo de estado SMS
' (*.MIF) cuando la instalación termine (con éxito o
' con un fallo).
'
' Tenga en cuenta que si gfSMS = False, no se escribirá ningún mensaje.
' Por lo tanto, para evitar que se registren algunos mensajes
' (p.e., sólo los mensajes de confirmación), establezca temporalmente
' gfSMS = False.
'
' Entrada: [strMsg] - El mensaje
'
' Normalmente, se llama a esta rutina en lugar de visualizar un
' MsgBox, strMsg es el mismo mensaje que debería aparecer
' en el MsgBox
'-----------------------------------------------------------
'
Sub LogSMSMsg(strMsg As String)
    If Not gfSMS Then Exit Sub
    '
    ' Agrega el mensaje. Tenga en cuenta que la longitud
    ' total no puede exceder de 255 caracteres, si es así
    ' se producirá un truncamiento.
    '
    gstrSMSDescription = Left(gstrSMSDescription & strMsg, MAX_SMS_DESCRIP)
End Sub

'-----------------------------------------------------------
' Función: MakePathAux
'
' Crea la ruta de directorio especificada.
'
' No hay interacción del usuario si se produce un error.
' Si se desea una interacción del usuario, use la función
'   MakePathAux().
'
' Entrada: [strDirName] - nombre de la ruta de directorio para crear
'
' Devuelve: True si tiene éxito, False si hay un error.
'-----------------------------------------------------------
'
Function MakePathAux(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Agrega un barra de separación final
    '
    If Right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Realiza un bucle y crea cada subdirectorio de la ruta separadamente.
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Comienza con al menos una barra invertida, p.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determina si el directorio existe
            Err = 0
            ChDir strPath
            If Err Then
                ' Tenemos que crear el directorio
                Err = 0
#If LOGGING Then
                NewAction gstrKEY_CREATEDIR, """" & strPath & """"
#End If
                MkDir strPath
#If LOGGING Then
                If Err Then
                    LogError ResolveResString(resMAKEDIR) & " " & strPath
                    AbortAction
                    GoTo Done
                Else
                    CommitAction
                End If
#End If
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function

'-----------------------------------------------------------
' Función: MsgError
'
' Establece el puntero del mouse con el predeterminado, llama a la
' función MsgBox de VB, registra el error (sólo 32 bits)
' y escribe el mensaje y la respuesta del usuario en el archivo de
' registro (sólo 32 bits)
'
' Entrada: [strMsg] - mensaje a visualizar
'          [intFlags] - indicadores de tipo de la función MsgBox
'          [strCaption] - título a usar en el cuadro de mensaje
'          [intLogType] (opcional) - El tipo del archivo de registro a crear.
'                       De forma predeterminada, crea una entrada de error. Usa
'                       la función MsgWarning() para crear una advertencia.
'                       Los tipos válidos son MSGERR_ERROR y MSGERR_WARNING
'
' Devuelve: El resultado de la función MsgBox
'-----------------------------------------------------------
'
Function MsgError(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String, Optional ByVal intLogType As Variant) As Integer
    Dim iRet As Integer
    
    iRet = MsgFunc(strMsg, intFlags, strCaption)
    MsgError = iRet
#If LOGGING Then
    ' Necesitamos registrar el error y decodificar la respuesta del usuario.
    Dim strID As String
    Dim strLogMsg As String

    Select Case iRet
        Case IDOK
            strID = ResolveResString(resLOG_IDOK)
        Case IDCANCEL
            strID = ResolveResString(resLOG_IDCANCEL)
        Case IDABORT
            strID = ResolveResString(resLOG_IDABORT)
        Case IDRETRY
            strID = ResolveResString(resLOG_IDRETRY)
        Case IDIGNORE
            strID = ResolveResString(resLOG_IDIGNORE)
        Case IDYES
            strID = ResolveResString(resLOG_IDYES)
        Case IDNO
            strID = ResolveResString(resLOG_IDNO)
        Case Else
            strID = ResolveResString(resLOG_IDUNKNOWN)
        'End Case
    End Select

    strLogMsg = strMsg & LF$ & "(" & ResolveResString(resLOG_USERRESPONDEDWITH, "|1", strID) & ")"
    If IsMissing(intLogType) Then
        intLogType = MSGERR_ERROR
    End If
    Select Case intLogType
        Case MSGERR_WARNING
            LogWarning strLogMsg
        Case MSGERR_ERROR
            LogError strLogMsg
        Case Else
            LogError strLogMsg
        'End Case
    End Select
#End If
End Function

'-----------------------------------------------------------
' Función: MsgFunc
'
' Establece el puntero del mouse con el predeterminado, llama a la
' función MsgBox de VB. Vea también MsgError.
'
' Entrada: [strMsg] - mensaje a visualizar
'          [intFlags] - indicadores de tipo de la función MsgBox
'          [strCaption] - título a usar en el cuadro de mensaje
' Devuelve: Resultado de la función MsgBox
'-----------------------------------------------------------
'
Function MsgFunc(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    Dim intOldPointer As Integer
  
    intOldPointer = Screen.MousePointer
    If gfNoUserInput Then
        MsgFunc = GetDefMsgBoxButton(intFlags)
        If gfSilent = True Then
            LogSilentMsg strMsg
        End If
        If gfSMS = True Then
            LogSMSMsg strMsg
            gfDontLogSMS = False
        End If
    Else
        Screen.MousePointer = gintMOUSE_DEFAULT
        MsgFunc = MsgBox(strMsg, intFlags, strCaption)
        Screen.MousePointer = intOldPointer
    End If
End Function

'-----------------------------------------------------------
' Función: MsgWarning
'
' Establece el puntero del mouse con el predeterminado, llama a la
' función MsgBox de VB, registra el error (sólo 32 bits)
' y escribe el mensaje y la respuesta del usuario en el archivo de
' registro (sólo 32 bits)
'
' Entrada: [strMsg] - mensaje a visualizar
'          [intFlags] - indicadores de tipo de la función MsgBox
'          [strCaption] - título a usar en el cuadro de mensaje
'
' Devuelve: El resultado de la función MsgBox
'-----------------------------------------------------------
'
Function MsgWarning(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    MsgWarning = MsgError(strMsg, intFlags, strCaption, MSGERR_WARNING)
End Function
'-----------------------------------------------------------
' Procedimiento: SetFormFont
'
' Se desplaza por los controles de un formulario especificado y
' establece las fuentes de las propiedades según el archivo de recursos.
'
' Entrada: [frm] - Formulario en el que se desean establecer las fuentes.
'-----------------------------------------------------------
'
Public Sub SetFormFont(frm As Form)
    Dim ctl As Control
    Dim fntSize As Integer
    Dim fntName As String
    Dim fntBold As Boolean
    '
    ' este procedimiento establecerá la fuente de cada control
    ' con una fuente del archivo de recursos
    '
    ' algunos controles pueden fallar, entonces se realizará un Resume Next...
    '
    On Error Resume Next
    '
    ' obtiene el nombre y el tamaño de la fuente desde el archivo de recursos
    '
    fntSize = CInt(LoadResString(resFONTSIZE))
    fntBold = (LoadResString(resFONTBOLD) = Format(True))
    fntName = LoadResString(resFONTNAME)
    '
    ' establece el nombre de la fuente del formulario (no establece el tamaño y si es negrita hasta más tarde)
    '
    frm.FontName = fntName
    If Err.Number <> 0 Then
        '
        ' Error al establecer la fuente. Probablemente
        ' la fuente no está instalada. Hay dos nombres
        ' de fuentes alternativos en el archivo de recursos.
        ' Pruebe con estos.
        '
        Err.Number = 0
        fntName = LoadResString(resFONTNAMEBACKUP1)
        frm.FontName = fntName
        If Err.Number <> 0 Then
            '
            ' Todavía no funciona. Una última prueba, pruebe
            ' con el segundo juego de fuentes. Si falla esta vez
            ' sólo podemos continuar y usar la fuente predeterminada.
            '
            '
            Err.Number = 0
            fntName = LoadResString(resFONTNAMEBACKUP2)
            frm.FontName = fntName
            Err.Number = 0
        End If
    End If
    '
    ' Ahora se establece FontSize y FontBold. Es necesario hacerlo
    ' después de establecer FontName ya que no todas las fuentes
    ' son compatibles con todos los tamaños, etc.
    '
    frm.FontSize = fntSize
    frm.FontBold = fntBold
    '
    ' recorre cada control e intenta establecer las propiedades de la fuente,
    ' puede fallar, pero entonces se activa un controlador de errores
    '
    For Each ctl In frm.Controls
        ctl.FontName = fntName
        ctl.FontSize = fntSize
        ctl.FontBold = fntBold
    Next
    '
    ' se restablece el controlador de errores
    '
    Set ctl = Nothing
    On Error GoTo 0
    Exit Sub
       
End Sub

'-----------------------------------------------------------
' Procedimiento: SetMousePtr
'
' Proporciona una forma de establecer el puntero del mouse sólo cuando
' cambia el estado del puntero. Para cada llamada a HOURGLASS, tendría
' que haber una llamada a DEFAULT. Es posible establecer otros
' tipos de punteros de mouse de forma explícita.
'
' Entrada: [intMousePtr] - tipo de puntero del mouse deseado
'-----------------------------------------------------------
'
Sub SetMousePtr(intMousePtr As Integer)
    Static intPtrState As Integer

    Select Case intMousePtr
        Case gintMOUSE_HOURGLASS
            intPtrState = intPtrState + 1
        Case gintMOUSE_DEFAULT
            intPtrState = intPtrState - 1
            If intPtrState < 0 Then
                intPtrState = 0
            End If
        Case Else
            Screen.MousePointer = intMousePtr
            Exit Sub
        'End Case
    End Select

    Screen.MousePointer = IIf(intPtrState > 0, gintMOUSE_HOURGLASS, gintMOUSE_DEFAULT)
End Sub

'-----------------------------------------------------------
' Función: StripTerminator
'
' Devuelve una cadena sin un terminador en cero. Es la forma típica
' en la que la API de Windows API devuelve una cadena.
'
' Entrada: [strString] - Cadena de la que eliminar el terminador
'
' Devuelve: El valor de la cadena que se pasa menos cualquier
'           cadena terminada en cero.
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'-----------------------------------------------------------
' Función: GetFileVersion
'
' Devuelve el número de versión interno del archivo
' especificado. Puede ser diferente que el número de versión
' visualizado por el cuadro de diálogo de Administrador de archivos.
' Es el número de versión que se muestra en la pantalla Detalles de
' archivos del Asistente para instalar de VB5. Es el número que usa
' la API de Windows VerInstallFile al comparar versiones de archivos.
'
' Entrada: [strFilename] - el archivo del que se desea saber el número de versión
'          [fIsRemoteServerSupportFile] - indica si es o no un archivo
'          de soporte de un componente ActiveX remoto (.VBR)
'          (sólo Edición empresarial). Si falta, se asume False.
'
' Devuelve: La cadena del número de versión, de lo contrario
'           gstrNULL
'-----------------------------------------------------------
'
Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
    Dim sVerInfo As VERINFO
    Dim strVer As String

    On Error GoTo GFVError

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If
    
    '
    'Obtiene la versión del archivo en una estructura VERINFO, y construye una cadena de
    'versión con los elementos apropiados.
    '
    If GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
        strVer = Format$(sVerInfo.nMSHi) & gstrDECIMAL & Format$(sVerInfo.nMSLo) & gstrDECIMAL
        strVer = strVer & Format$(sVerInfo.nLSHi) & gstrDECIMAL & Format$(sVerInfo.nLSLo)
        GetFileVersion = strVer
    Else
        GetFileVersion = gstrNULL
    End If
    
    Exit Function
    
GFVError:
    GetFileVersion = gstrNULL
    Err = 0
End Function

'-----------------------------------------------------------
' Función: GetFileVerStruct
'
' Obtiene la información de versión de archivo en una
' variable VERINFO TYPE
'
' Entrada: [strFilename] - nombre del archivo del que obtener la versión
'          [fIsRemoteServerSupportFile] - indica si este archivo es un
'          archivo de soporte de un componente ActiveX remoto (.VBR)
'          (Sólo Edición empresarial). Si falta, se asume False.
' Salida: [sVerInfo] - Tipo VERINFO para completar con la información de la versión
'
' Devuelve: True si se encuentra la información de la versión, False en caso contrario
'-----------------------------------------------------------
'
Function GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
    Const strFIXEDFILEINFO$ = "\"

    Dim lVerSize As Long
    Dim lVerHandle As Long
    Dim lpBufPtr As Long
    Dim byteVerData() As Byte
    Dim fFoundVer As Boolean

    GetFileVerStruct = False
    fFoundVer = False

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If
    
    If fIsRemoteServerSupportFile Then
        GetFileVerStruct = GetRemoteSupportFileVerStruct(strFilename, sVerInfo)
        fFoundVer = True
    Else
        '
        'Obtiene el tamaño de la información de versión del archivo, asigna un búfer y obtiene
        'la información de versión. Después, consultamos el archivo Fixed por partes, donde está
        'la versión interna del archivo que usa la llamada VerInstallFile de la API. Después se
        'copia la información del archivo en una estructura VERINFO.
        '
        lVerSize = GetFileVersionInfoSize(strFilename, lVerHandle)
        If lVerSize > 0 Then
            ReDim byteVerData(lVerSize)
            If GetFileVersionInfo(strFilename, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pasa la matriz byteVerData al primer elemento mediante referencia)
                If VerQueryValue(byteVerData(0), strFIXEDFILEINFO & "", lpBufPtr, lVerSize) <> 0 Then
                    lmemcpy sVerInfo, lpBufPtr, lVerSize
                    fFoundVer = True
                    GetFileVerStruct = True
                End If
            End If
        End If
    End If
    
    If Not fFoundVer Then
        '
        ' No hemos podido obtener la información de versión del archivo.
        ' Es posible que sea un archivo de dependencias.
        '
        If UCase(Extension(strFilename)) = gstrEXT_DEP Then
            GetFileVerStruct = GetDepFileVerStruct(strFilename, sVerInfo)
        End If
    End If
End Function
'-----------------------------------------------------------
' Función: GetDepFileVerStruct
'
' Obtiene la información de versión de un archivo mediante su
' archivo de dependencias (*.dep). Como los archivos no tiene una marca de
' versión de Windows, pero tiene una marca interna que se puede
' consultar.
'
' Entrada: [strFilename] - Nombre del archivo de dependencias del que obtener la información
' Salida:  [sVerInfo] - Tipo de VERINFO para rellenar con la información de versión
'
' Devuelve: True si se encuentra la información de versión, False en caso contrario
'-----------------------------------------------------------
'
Function GetDepFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    GetDepFileVerStruct = False
    
    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile
    
    ' Recorre cada línea, buscando la clave
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If Left$(strLine, cchVersionKey) = strVersionKey Then
            ' Hemos encontrado la clave de la versión. Se copia todo lo que está después del signo igual
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Procesa y almacena la información de la versión
            PackVerInfo strVersion, sVerInfo

            GetDepFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetDepFileVerStruct = False
End Function

'-----------------------------------------------------------
' Función: GetRemoteSupportFileVerStruct
'
' Obtiene la información de versión del archivo de soporte de un
' componente Active X remoto en una variable VERINFO TYPE (sólo Edición
' Empresarial). Los archivos no tiene una marca de Windows
' paro tienen una marca de versión interna, gracias a la que
' podemos buscar.
'
' Entrada: [strFilename] - nombre del archivo del que obtener la versión
' Salida:  [sVerInfo] - tipo VERINFO para completar con la información de la versión
'
' Devuelve: True si se encuentra información de la versión, False en caso contrario
'-----------------------------------------------------------
'
Function GetRemoteSupportFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile
    
    ' Recorre cada línea, buscando la clave
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If Left$(strLine, cchVersionKey) = strVersionKey Then
            ' Hemos encontrado la clave de la versión. Se copia todo lo que está después del signo igual
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Procesa y almacena la información de la versión
            PackVerInfo strVersion, sVerInfo

            'Convierte el formato 1.2.3 del archivo .VBR a
            '1.2.0.3, qué es realmente el que queremos
            sVerInfo.nLSLo = sVerInfo.nLSHi
            sVerInfo.nLSHi = 0
            
            GetRemoteSupportFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetRemoteSupportFileVerStruct = False
End Function
'-----------------------------------------------------------
' Función: GetWindowsDir
'
' Llama a la API de Windows para obtener el directorio de Windows
' y comprueba que existe un separador de directorios inicial
'
' Devuelve: El directorio de Windows
'-----------------------------------------------------------
'
Function GetWindowsDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Obtiene el directorio de Windows y recorta el búfer con la longitud exacta
    'que se devuelve agregando un carácter separador (barra invertida) si la API no lo hace
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        GetWindowsDir = strBuf
    Else
        GetWindowsDir = gstrNULL
    End If
End Function

'-----------------------------------------------------------
' Función: UCase16
'
' Devuelve la conversión a mayúsculas de una cadena sobre
'   16 bits, o devuelve una copia sin modificar de la
'   cadena sobre 32 bits.
'
' Entrada: [str] - Cadena sobre la que copiar a mayúsculas
'
'-----------------------------------------------------------
'
Function UCase16(ByVal str As String)
    UCase16 = str
End Function

'-----------------------------------------------------------
' Función: ExtractFilenameItem
'
' Extrae de una cadena un nombre de archivo con o sin comillas.
'
' Entrada: [str] - cadena a procesar .
'          [intAnchor] - índice de str dónde comienza el nombre del archivo.
'             EL nombre del archivo continúa hasta el final de la cadena
'             o hasta la siguiente coma en la cadena, o, si el nombre
'             de archivo está entre comillas, hasta la siguiente
'             doble comilla.
' Salida:  Devuelve el nombre del archivo, sin comillas.
'      [intAnchor] se establece con una coma, o con el carácter del
'             final de la cadena
'      [fErr] se establece a True si se descubre un error al procesar
'
'-----------------------------------------------------------
'
Function strExtractFilenameItem(ByVal str As String, intAnchor As Integer, fErr As Boolean) As String
    While Mid$(str, intAnchor, 1) = " "
        intAnchor = intAnchor + 1
    Wend
    
    Dim iEndFilenamePos As Integer
    Dim strFilename As String
    If Mid$(str, intAnchor, 1) = """" Then
        ' El nombre del archivo está encerrado entre comillas
        iEndFilenamePos = InStr(intAnchor + 1, str, """") ' Encontrar las comillas
        If iEndFilenamePos > 0 Then
            strFilename = Mid$(str, intAnchor + 1, iEndFilenamePos - 1 - intAnchor)
            intAnchor = iEndFilenamePos + 1
            While Mid$(str, intAnchor, 1) = " "
                intAnchor = intAnchor + 1
            Wend
            If (Mid$(str, intAnchor, 1) <> gstrCOMMA) And (Mid$(str, intAnchor, 1) <> "") Then
                fErr = True
                Exit Function
            End If
        Else
            fErr = True
            Exit Function
        End If
    Else
        ' El nombre del archivo continúa hasta la siguiente coma o hasta el final de la cadena
        Dim iCommaPos As Integer
        
        iCommaPos = InStr(intAnchor, str, gstrCOMMA)
        If iCommaPos = 0 Then
            iCommaPos = Len(str) + 1
        End If
        iEndFilenamePos = iCommaPos
        
        strFilename = Mid$(str, intAnchor, iEndFilenamePos - intAnchor)
        intAnchor = iCommaPos
    End If
    
    strFilename = Trim$(strFilename)
    If strFilename = "" Then
        fErr = True
        Exit Function
    End If
    
    fErr = False
    strExtractFilenameItem = strFilename
End Function

'-----------------------------------------------------------
' Función: Extension
'
' Extrae la parte de la extensión en un nombre de archivo o de ruta
'
' Entrada: [strFilename] - archivo o ruta del que obtener la extensión
'
' Devuelve: La extensión si existe, si no es así devuelve gstrNULL
'-----------------------------------------------------------
'
Function Extension(ByVal strFilename As String) As String
    Dim intPos As Integer

    Extension = gstrNULL

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case gstrSEP_EXT
                Extension = Mid$(strFilename, intPos + 1)
                Exit Do
            Case gstrSEP_DIR, gstrSEP_DIRALT
                Exit Do
            'End Case
        End Select

        intPos = intPos - 1
    Loop
End Function

'-----------------------------------------------------------
' Procedimiento: PackVerInfo
'
' Procesa una cadena de número de versión con formato
' x[.x[.x[.x]]] y asigna los número extraídos con los
' elementos apropiados de una variable de tipo VERINFO.
' Ejemplos de cadenas de versión válidas son '3.11.0.102',
' '3.11', '3', etc.
'
' Entrada: [strVersion] - cadena de número de versión
'
' Salida:  [sVerInfo] - variable de tipo VERINFO cuyos elementos
'                   se asignan a los números adecuados
'                   de la cadena de número de versión
'-----------------------------------------------------------
'
Sub PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
    Dim intOffset As Integer
    Dim intAnchor As Integer

    On Error GoTo PVIError

    intOffset = InStr(strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
    Else
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    End If

    Exit Sub

PVIError:
    sVerInfo.nMSHi = 0
PVIMSLo:
    sVerInfo.nMSLo = 0
PVILSHi:
    sVerInfo.nLSHi = 0
PVILSLo:
    sVerInfo.nLSLo = 0
End Sub

Public Function strQuoteString(strUnQuotedString As String, Optional vForce As Variant, Optional vTrim As Variant)
'
' Esta rutina agrega comillas a una cadena que no las tiene, de forma predeterminada. Si la cadena ya tiene comillas
' la devuelve sin hacer ningún cambio a menos que vForce sea True (el valor predeterminado de vForce es False) excepto que
' se eliminarán los espacios en blanco antes y después de las comillas aunque vTrim sea False. Si la cadena contiene espacios
' en blanco iniciales o finales se eliminarán amenos que vTrim sea False (el valor predeterminado de vTrim es True).
'
    Dim strQuotedString As String
    
    If IsMissing(vForce) Then
        vForce = False
    End If
    If IsMissing(vTrim) Then
        vTrim = True
    End If
    
    strQuotedString = strUnQuotedString
    '
    ' Recorta la cadena si es necesario
    '
    If vTrim = True Then
        strQuotedString = Trim(strQuotedString)
    End If
    '
    ' Comprueba que la cadena ya tiene comillas
    '
    If vForce = False Then
        If (Left(strQuotedString, 1) = gstrQUOTE) And (Right(strQuotedString, 1) = gstrQUOTE) Then
            '
            ' La cadena ya tiene comillas. Está hecho.
            '
            GoTo DoneQuoteString
        End If
    End If
    '
    ' Agrega las comillas
    '
    strQuotedString = gstrQUOTE & strQuotedString & gstrQUOTE
DoneQuoteString:
    strQuoteString = strQuotedString
End Function
Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' Esta rutina comprueba que strQuotedString está encerrada entre
' comillas, y si es así, las elimina.
'
    strQuotedString = Trim(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And Right$(strQuotedString, 1) = gstrQUOTE Then
        '
        ' Si tiene comillas, quita las comillas.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    strUnQuoteString = strQuotedString
End Function
Public Function fCheckFNLength(strFilename As String) As Boolean
'
' Esta rutina comprueba que la longitud del nombre de archivo strFilename es válida.
' Con NT (Intel) y Win95 puede ser de hasta 259 (gintMAX_PATH_LEN-1) caracteres de
' longitud. Esta longitud debe incluir la unidad, ruta, nombre de archivo, argumentos
' de la línea de comandos y comillas (si la cadena las tiene).
'
    fCheckFNLength = (Len(strFilename) < gintMAX_PATH_LEN)
End Function
Public Function intGetNextFldOffset(ByVal intAnchor As Integer, strList As String, strDelimit As String, Optional CompareType As Variant) As Integer
'
' Esta rutina lee de una lista separada strDelimit, strList, y coloque el siguiente
' elemento de la lista a continuación de intAnchor. Básicamente encuentra la siguiente
' aparición de strDelimit que no tiene comillas. Si strDelimit no encuentra la rutina
' devuelve 0. Tenga en cuenta que intAnchor debe estar fuera de las comillas
' o esta rutina devolverá resultados incorrectos.
'
' strDelimit es normalmente una coma.
'
' Si hay un error esta rutina devuelve -1.
'
    Dim intQuote As Integer
    Dim intDelimit As Integer
    
    Const CompareBinary = 0
    Const CompareText = 1

    If IsMissing(CompareType) Then
        CompareType = CompareText
    End If
    
    If intAnchor = 0 Then intAnchor = 1
    
    intQuote = InStr(intAnchor, strList, gstrQUOTE, CompareType)
    intDelimit = InStr(intAnchor, strList, strDelimit, CompareType)
    
    If (intQuote > intDelimit) Or (intQuote = 0) Then
        '
        ' El siguiente delimitador no está entre comillas. Por lo tanto,
        ' hemos encontrado lo que estamos buscando. Tenga en cuenta que el
        ' caso en el que no hay delimitadores también se trata aquí.
        '
        GoTo DoneGetNextFldOffset
    ElseIf intQuote < intDelimit Then
        '
        ' Aparece una comilla antes del siguiente delimitador. Significa
        ' que podemos estar entre comillas. Todavía es necesario comprobar
        ' si la comilla de cierre está después del delimitador o no.
        '
        intAnchor = intQuote + 1
        intQuote = InStr(intAnchor, strList, gstrQUOTE, CompareType)
        If (intQuote > intDelimit) Then
            '
            ' El delimitador está entre comillas. Por lo tanto, se ignora.
            ' El siguiente delimitador después de la comilla de cierre debe estar
            ' fuera de las comillas, de lo contrario tenemos un archivo dañado.
            '
            intAnchor = intQuote + 1
            intDelimit = InStr(intAnchor, strList, strDelimit, CompareType)
            '
            ' Comprobación de seguridad. Comprueba que no hay más comillas
            ' después del delimitador que hemos encontrado.
            '
            If intDelimit > 0 Then
                intQuote = InStr(intAnchor, strList, gstrQUOTE, CompareType)
                If (intQuote > 0) And (intQuote < intDelimit) Then
                    '
                    ' Algo va mal. Hemos encontrado una comilla perdida.
                    ' Significa que la cadena probablemente está dañada.
                    '
                    intDelimit = -1 ' Error
                End If
            End If
        End If
    End If
DoneGetNextFldOffset:
    intGetNextFldOffset = intDelimit
End Function

