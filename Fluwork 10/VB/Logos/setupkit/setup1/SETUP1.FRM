VERSION 5.00
Begin VB.Form frmSetup1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "VB5 Setup Toolkit"
   ClientHeight    =   1770
   ClientLeft      =   225
   ClientTop       =   1590
   ClientWidth     =   7950
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "setup1.frx":0000
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Label lblModify 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"setup1.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   7800
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDDE 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esta etiqueta se usar para una conexión DDE con el Administrador de programas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   1515
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "frmSetup1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'
' Imposible situar esta cadena en un recurso ya que indica el fallo al cargar recursos, es necesario traducirla de forma separada
'
Const mstrRESOURCELOADFAIL$ = "Error al inicializar los recursos de cadenas que usa el programa de instalación."

'-----------------------------------------------------------
' Procedimiento: DrawBackGround
'
' Dibuja la pantalla azul e imprime el título sombreado de
' la aplicación de instalación
'-----------------------------------------------------------
'
Private Sub DrawBackGround()
    Const intBLUESTART% = 255
    Const intBLUEEND% = 0
    Const intBANDHEIGHT% = 2
    Const intSHADOWSTART% = 8
    Const intSHADOWCOLOR% = 0
    Const intTEXTSTART% = 4
    Const intTEXTCOLOR% = 15

    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer

    '
    'Obtiene los valores del sistema para anchura y altura
    '
    intFormHeight = ScaleHeight
    intFormWidth = ScaleWidth

    '
    'Calcula el tamaño del paso y el valor de azul inicial
    '
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
    sngBlueCur = intBLUESTART

    '
    'Dibuja la pantalla azul
    '
    For intY = 0 To intFormHeight Step intBANDHEIGHT
        Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
        sngBlueCur = sngBlueCur + sngBlueStep
    Next intY

    '
    'Imprime con sombra el nombre de la aplicación
    '
    CurrentX = intSHADOWSTART
    CurrentY = intSHADOWSTART
    ForeColor = QBColor(intSHADOWCOLOR)
    Print Caption
    CurrentX = intTEXTSTART
    CurrentY = intTEXTSTART
    ForeColor = QBColor(intTEXTCOLOR)
    Print Caption
End Sub
Private Sub Form_Load()
'
' La mayor parte del trabajo de Setup1 tiene lugar en Form_Load()
' y se basa en la información que aparecen en el archivo
' SETUP.LST. Para personalizar la funcionalidad de Setup1,
' probablemente querrá modificar el archivo SETUP.LST.
' En particular, la información que indica los archivos que
' está instalando se almacena en SETUP.LST. Las únicas excepciones
' son los archivos de Automatización remota RacMgr32.Exe
' y AutMgr32.Exe que necesitan un tratamiento especial
' para instalar sus iconos en un grupo de programas
' especial.
'
' También se puede personalizar editando el código
' que sigue a Form_Load o en otras partes del programa.
' El código que se puede personalizar está
' documentado con sugerencias y ejemplos de código.
'
    Const strINI_FILES$ = "Files"                           'sección predeterminada para instalar
    Const strEXT_GRP$ = "GRP"                               'extensión del grupo de programas
    Const SW_HIDE = 0

    Dim strGroupName As String                              'Nombre del grupo de programas
    Dim sFile As FILEINFO                                   'primeros archivos= información de línea
    gfRegDAO = False
    
    On Error GoTo MainError

    SetFormFont Me
    Me.Font.Size = 24
    Me.Font.Bold = True
    '
    'Inicializa las variables de espaciado de línea para los mensajes, etc.
    '
    LF$ = Chr$(10)
    LS$ = LF$ & LF$
    CRLF = Chr$(13) & Chr$(10)
    '
    'Inicializa los recursos de cadena que usan las variables globales y los controles y formularios
    '
    GetStrings
    
    '
    'Obtiene los directorios Windows y Windows\System
    '
    gstrWinDir = GetWindowsDir()
    gstrWinSysDir = GetWindowsSysDir()

    '
    ' Si el directorio Windows System es un subdirectorio del directorio de
    ' Windows, el lugar correcto para la instalación de los archivos
    ' especificados en la setup.lst como $(WinSysDest) es siempre el
    ' directorio Windows \System. Si el directorio Windows\System
    ' no es un subdirectorio del directorio de Windows,
    ' el usuario está ejecutando una versión compartida de Windows. En
    ' este caso, si el usuario no tiene derechos de acceso al directorio
    ' de sistema compartido, cambiamos el destino de los archivos de
    ' sistema al directorio de Windows.
    '
    If InStr(gstrWinSysDir, gstrWinDir) = 0 Then
        If WriteAccess(gstrWinSysDir) = False Then
            gstrWinSysDir = gstrWinDir
        End If
    End If

    '
    ' Es necesario procesar los argumentos de la línea de comandos tan pronto
    ' como sea posible, ya que sin ellos es imposible llamar
    ' a la aplicación que elimina los archivos al cancelar el proceso
    ' de instalación.
    '
    ProcessCommandLine Command$, gfSilent, gstrSilentLog, gfSMS, gstrMIFFile, gstrSrcPath, gstrAppRemovalLog, gstrAppRemovalEXE
    gfNoUserInput = (gfSilent Or gfSMS)
    
    AddDirSep gstrSrcPath

    '
    ' El programa de instalación (SETUP.EXE) copia SETUP1.EXE y SETUP.LST en el
    ' directorio de windows del usuario final. La información necesaria para la instalación
    ' como los indicadores de instalación y la información de los archivos se leen de la copia de SETUP.LST
    ' que se encuentra en ese directorio.
    '
    gstrSetupInfoFile = gstrWinDir & gstrFILE_SETUP
    
    gstrAppName = ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_APPNAME)
    If gstrAppName = gstrNULL Then
        MsgError ResolveResString(resNOSETUPLST), MB_OK Or MB_ICONSTOP, gstrSETMSG
        gstrTitle = ResolveResString(resSETUP, "|1", gstrAppName)
        ExitSetup Me, gintRET_FATAL
    End If
    
    gstrTitle = ResolveResString(resSETUP, "|1", gstrAppName)
    If gfSilent Then LogSilentMsg gstrTitle & CRLF
    '
    ' Éste es el nombre predeterminado del grupo en el que se instalan los iconos.
    '
    gstrDefGroup = ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_DEFGROUP)
    If gstrDefGroup = gstrNULL Then
        gstrDefGroup = gstrAppName
    End If

    '
    ' Muestra la pantalla de instalación con el fondo azul tan pronto como se tiene el título
    '
    ShowMainForm

    '
    ' Muestra el cuadro de diálogo de bienvenida
    '
    ShowWelcomeForm

    '
    ' Obtiene el nombre del archivo ejecutable de la aplicación. Este nombre se agregará
    ' al Administrador de programas o al Explorador. Sólo hay que introducir el nombre del EXE
    ' en esta clave del archivo de información de instalación (SETUP.LST).
    '
    gstrAppExe = ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_APPEXE)
    '
    ' Obtiene el nombre de la aplicación para desinstalación. En la mayor parte de los casos,
    ' será el mismo que gstrAppExe excepto cuando la aplicación sea un componente compartido.
    ' gstrAppToUninstall se usa por la desinstalación de SMS para determinar la ubicación de
    ' los archivos de desinstalación.
    '
    gstrAppToUninstall = ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_APPTOUNINSTALL)
    
    '
    ' Si se establece este indicador, el directorio de destino especificado se usa
    ' sin preguntar y el usuario no puede cambiarlo.
    ' Esto se aplica para instalar un .EXE o una .DLL como un componente en lugar
    ' de una aplicación en un directorio de aplicación. En este caso,
    ' no tiene sentido tener un directorio de aplicación.
    '
    If ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_FORCEUSEDEFDEST) = "1" Then
        gfForceUseDefDest = True
    End If
    
    '
    ' Se lee el directorio de destino predeterminado. Si el nombre especificado está en conflicto
    ' con el nombre de un archivo, se pide un nuevo directorio predeterminado
    '
    gstrDestDir = ResolveDestDir(ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_APPDIR))
    While FileExists(gstrDestDir) = True Or gstrDestDir = gstrNULL
        If MsgError(ResolveResString(resBADDEFDIR), MB_OKCANCEL Or MB_ICONQUESTION, gstrSETMSG) = IDCANCEL Then
            ExitSetup Me, gintRET_FATAL
        End If
        
        If gfNoUserInput = True Then
            ExitSetup Me, gintRET_FATAL
        Else
            ShowPathDialog gstrDIR_DEST
        End If
    Wend

    '
    ' Asegura una barra invertida en el inicio del directorio de destino
    '
    AddDirSep gstrDestDir

    Do
        '
        ' Muestra el botón de instalación y el directorio predeterminado. El usuario puede
        ' cambiar aquí el directorio de destino.
        '
        ShowBeginForm

        '
        ' Éste es un buen sitio para visualizar un cuadro de diálogo de opciones, que permita
        ' al usuario seleccionar opciones de instalación: ejemplos, documentación, archivos de Ayuda, etc.
        ' Los resultados de este cuadro de diálogo se comprobarían en el siguiente bucle
        '
        'ShowOptionsDialog (Función que deberá escribir para comprobar las opciones, etc.)
        '

        '
        ' Inicializa la "tabla" de unidades y la matriz de espacio en disco
        '
        InitDiskInfo

        SetMousePtr gintMOUSE_HOURGLASS
        ShowStaticMessageDialog ResolveResString(resDISKSPACE)

        '
        ' Para cada sección de SETUP.LST que se instala, se llama a CalcDiskSpace
        ' con el nombre de la sección
        '
        CalcDiskSpace strINI_FILES
        'CalcDiskSpace "MySection"
        'CalcDiskSpace "MyOtherSection"
        '
        ' Si crea un cuadro de diálogo de opciones, necesita comprobar los resultados
        ' para calcular el espacio en disco necesario (si se instala con
        ' opciones)
        '
        'If chkInstallSamples.Value = TRUE then
        '    CalcDiskSpace "Samples"
        'End If
        '

        HideStaticMessageDialog
        SetMousePtr gintMOUSE_DEFAULT

    '
    ' Después de completar las llamadas a CalcDiskSpace, se llama a CheckDiskSpace para
    ' comprobar los resultados y mostrar el formulario de advertencia (si fuese necesario). Si el usuario
    ' desea probar en otro directorio de destino (o borrar y volver a intentarlo) entonces
    ' CheckDiskSpace devolverá False
    '
    Loop While CheckDiskSpace() = False
    
    '
    ' Comienza iniciando el archivo de registro de instalación (se usará para eliminar la aplicación)
    '
    EnableLogging gstrAppRemovalLog
    '
    ' Continúa y obliga a crear el directorio de la aplicación,
    ' ya que el archivo de registro para desinstalar la aplicación se copiará allí.
    '
    MakePath gstrDestDir, False 'EL usuario puede que no ignore los errores
    
    '
    ' Crea el grupo de programa principal si se necesita uno.
    '
    Const fDefCreateGroupUnderWin95 = False
    '
    ' Si se establece fDefCreateGroupUnderWin95 a False (predeterminado), no se creará
    ' un grupo de programa en Win95 a menos que sea absolutamente necesario.
    '
    ' Por defecto en Windows 95, no se creará un grupo, y el único icono de
    ' programa se situará directamente en el menú
    ' Inicio>Programas (aunque hay otro, se crea con iconos definidos por el usuario
    '
    Dim cIcons As Integer            ' Contador de los iconos necesarios.
    Dim fAdditionalIcons As Boolean
    '
    ' Hay dos razones por las que instalar más iconos que los de
    ' simplemente el programa ejecutable:
    '   1) Si estamos instalando iconos definidos por el usuario (personalizando este procedimiento
    '      o personalizando SETUP.LST)
    '   2) Si estamos creando un icono de un programa de desinstalación (al ejecutar con NT)
    '
    ' Lee el archivo SETUP.LST y calcula los iconos necesarios.
    '
    cIcons = CountIcons(strINI_FILES)
    '
    ' Hace lo mismo en otras secciones de SETUP.LST si las ha agregado.
    '
    'cIcons = cIcons + CountIcons("MySection")
    'cIcons = cIcons + CountIcons("MyOtherSection")
    
    '
    ' Si ha modificado este procedimiento para instalar más iconos, compruebe que establece
    ' esta variable a True, para que se cree un grupo de programas para los iconos.
    ' Actualmente, es True sólo si se encuentra más de una entrada en la sección "Files"
    ' de SETUP.LST que especifique un icono o si es NT 3.51.
    '
    fAdditionalIcons = False
    fAdditionalIcons = fAdditionalIcons Or (cIcons > 1)
    fAdditionalIcons = fAdditionalIcons Or (Not TreatAsWin95())
    '
    ' La siguiente variable determina si se crea un grupo de programas
    ' para los iconos. Se controla con fNoGroupUnderWin95,
    ' fAdditionalIcons y FTreatAsWin95().
    '
    Dim fCreateGroup As Boolean
    If TreatAsWin95() Then
        '
        ' Sólo Win95:
        ' Se crea un grupo de programas sólo si tenemos iconos adicionales
        ' para el ejecutable de la aplicación (si hay), o si se ha establecido fDefCreateGroupUnderWin95
        ' a True para omitir este comportamiento predeterminado.
        '
        fCreateGroup = fDefCreateGroupUnderWin95 Or fAdditionalIcons
    Else
        '
        ' Sólo Win32 NT:
        ' Siempre se debe crear un grupo del Administrador de programas
        ' ya que siempre se crea un icono para la aplicación de desinstalación.
        '
        fCreateGroup = True
    End If
    
    strGroupName = ""
    If fCreateGroup Then
        strGroupName = frmGroup.GroupName(frmSetup1, gstrDefGroup)
        fMainGroupWasCreated = True
    End If

    '
    ' Muestra el formulario de copia y establece el indicador de porcentaje a cero
    '
    SetMousePtr gintMOUSE_HOURGLASS
    ShowCopyDialog
    UpdateStatus frmCopy.picStatus, 0, True

    '
    ' Siempre se comienza con el disco nº1
    '
    gintCurrentDisk = 1
    '
    ' Para cada sección de SETUP.LST que sea necesario instalar, llame a CopySection
    ' con el nombre de la sección
    '
    
    CopySection strINI_FILES
    'CopySection "MySection"
    'CopySection "MyOtherSection"
        
    '
    ' Si crea un cuadro de diálogo de opciones, necesita comprobar los resultados aquí
    ' para determinar si se copian los archivos a una sección en particular.
    '
    'If chkInstallSamples.Value = TRUE then
    '    CopySection "Samples"
    'End If
    '

    UpdateStatus frmCopy.picStatus, 1, True
    
    HideCopyDialog

    '
    ' Si instalamos AXDIST.EXE, necesitamos ejecutarlo
    ' para que instale los archivos adicionales que contiene.
    '
    If gfAXDist = True Then
        '
        'Ejecuta el shell de forma síncrona y la utilidad con los modificadores correctos
        '
        If FileExists(gstrAXDISTInstallPath) Then
            FSyncShell gstrAXDISTInstallPath, SW_HIDE
        End If
    End If
    '
    ' Si WINt351.EXE está instalado, necesitamos ejecutarlo
    ' para que instale los archivos adicionales que contiene.
    '
    If gfWINt351 = True Then
        '
        'Ejecuta el shell de forma síncrona y la utilidad con los modificadores correctos
        '
        If FileExists(gstrWINt351InstallPath) Then
            FSyncShell gstrWINt351InstallPath, SW_HIDE
        End If
    End If
    '
    ' Ahora se realizan las actualizaciones 'invisibles' que sean necesarias
    '
    SetMousePtr gintMOUSE_DEFAULT
    ShowStaticMessageDialog ResolveResString(resUPDATING)

    '
    ' Registra todos los archivos que se han guardado en la matriz de registro.
    ' CopySection de API agrega una entrada de registro (si es necesaria) al copiar un archivo.
    '
    RegisterFiles
    
    '
    ' Registra todas las licencias que aparecen en las sección [Licenses] de
    ' Setup.lst.
    '
    RegisterLicenses
    
    '
    ' Si están instalados archivos DAO, necesitamos agregar algunas claves especiales
    ' al Registro para permitir que se funcionen con campos de bases de datos OLE.
    '
    '
    If gfRegDAO = True Then
        RegisterDAO
    End If
    '
    ' Crea iconos de programa (o vínculos, por ejemplo accesos directos).
    '
    If (fMainGroupWasCreated = True) Or ((cIcons > 0) And TreatAsWin95()) Then
        ShowStaticMessageDialog ResolveResString(resPROGMAN)
        CreateIcons strINI_FILES, strGroupName
        '
        ' Hace lo mismo para las otras secciones de SETUP.LST si las ha agregado.
        '
        'CreateIcons "MySection", strGroupName
        'CreateIcons "MyOtherSection", strGroupName
        '
    End If
    '
    ' Si necesita crear más iconos, inserte código aquí y compruebe que ha establecido
    ' la variable fAdditionalIcons a True
    '
    ' If Not fAdditionalIcons Then
    '   MsgError "Internal Setup Customization Error: fAdditionalIcons = False", vbOKOnly Or vbExclamation, gstrTitle
    '   ExitSetup Me, gintRET_FATAL
    ' End If
    ' CreateOSLink frmSetup1, strGroupName, gsDest.strAppDir & "My Exe 1.exe", "My Exe 1 command-line arguments", "My Exe 1"
    ' CreateOSLink frmSetup1, strGroupName, gsDest.strAppDir & "My Exe 2.exe", "My Exe 2 command-line arguments", "My Exe 2"
    '



    '
    ' Crea un grupo de programas separado e iconos para el Administrador de
    ' automatización remota y el Administrador de automatización, si se
    ' ha instalado.
    ' Este grupo de programas está separado completamente del que crea el programa de
    ' la aplicación (si existe), ya que se compartirá por todas las aplicaciones
    ' de VB instaladas.
    '
    ' NOTA: Esto no instalará iconos adicionales. Se controla
    ' NOTA: después de crear los iconos de Automatización remota.
    '
    ShowStaticMessageDialog ResolveResString(resPROGMAN)
    If gsDest.strAUTMGR32 <> "" Or gsDest.strRACMGR32 <> "" Then
        'Al menos se ha instalado uno de estos programas. Continúa
        'y crea el grupo de programas.
        Dim strRemAutGroupName As String
        
        strRemAutGroupName = ResolveResString(resREMAUTGROUPNAME)
        '
        ' Crea el grupo de iconos de Automatización remota. Tenga en cuenta
        ' que el usuario no puede elegir el nombre de este grupo, no hay forma
        ' de solucionar un error si ocurriese. Por lo tanto,
        ' fCreateOSProgramGroup cancelaría la instalación, sin volver si
        ' ocurriese un error.
        '
        fCreateOSProgramGroup frmSetup1, strRemAutGroupName, False, False

        ' Ahora se crean los iconos de AUTMGR32.EXE y RACMGR32.EXE
        If gsDest.strRACMGR32 <> "" Then
            CreateOSLink frmSetup1, strRemAutGroupName, gsDest.strRACMGR32, "", ResolveResString(resRACMGR32ICON), False
        End If
        If gsDest.strAUTMGR32 <> "" Then
            CreateOSLink frmSetup1, strRemAutGroupName, gsDest.strAUTMGR32, "", ResolveResString(resAUTMGR32ICON), False
        End If
    End If

    '
    ' Registra la ruta de la aplicación
    '
    If gstrAppExe <> "" Then
        Dim strPerAppPath As String
        strPerAppPath = ReadIniFile(gstrSetupInfoFile, gstrINI_SETUP, gstrINI_APPPATH)
        AddPerAppPath gstrAppExe, gsDest.strAppDir, strPerAppPath
    End If

ExitSetup:
    HideStaticMessageDialog
    RestoreProgMan
    If fWithinAction() Then
        'Por ahora, todas las acciones de registro deberían ser cortadas o realizadas.
        MsgError ResolveResString(resSTILLWITHINACTION), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup Me, gintRET_FATAL
    End If
    MoveAppRemovalFiles strGroupName
    '
    ' Si se instala AXDIST.EXE, es necesario eliminarlo.
    ' Es un archivo de extracción automática que ya está extraído.
    '
    If FileExists(gstrAXDISTInstallPath) Then
        Kill gstrAXDISTInstallPath
    End If
    '
    ' Si se instala WINt351.EXE, es necesario eliminarlo.
    ' Es un archivo de extracción automática que ya está extraído.
    '
    If FileExists(gstrWINt351InstallPath) Then
        Kill gstrWINt351InstallPath
    End If
    
    ExitSetup Me, gintRET_FINISHEDSUCCESS

MainError:
    Dim iRet As Integer
    iRet = MsgError(Error$ & LS$ & ResolveResString(resUNEXPECTED), MB_RETRYCANCEL Or MB_ICONEXCLAMATION, gstrTitle)
    If gfNoUserInput Then iRet = IDCANCEL
    Select Case iRet
        Case IDRETRY
            Resume
        Case IDCANCEL
            ExitSetup Me, gintRET_ABORT
            Resume
        'End Case
    End Select
End Sub

'-----------------------------------------------------------
' Procedimiento: HideCopyDialog
'
' Descarga el formulario de estado de la copia de archivos
'-----------------------------------------------------------
'
Private Sub HideCopyDialog()
    Unload frmCopy
End Sub

'-----------------------------------------------------------
' Procedimiento: HideStaticMessageDialog
'
' Descarga el formulario de mensajes de la instalación
'-----------------------------------------------------------
'
Private Sub HideStaticMessageDialog()
    Unload frmMessage
End Sub

'-----------------------------------------------------------
' Procedimiento: ShowBeginForm
'
' Muestra el formulario de inicio de la instalación
'-----------------------------------------------------------
'
Private Sub ShowBeginForm()
    If gfNoUserInput Then
        If IsValidDestDir(gstrDestDir) = False Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    Else
        frmBegin.Show 1
    End If
End Sub

'-----------------------------------------------------------
' Procedimiento: ShowCopyDialog
'
' Muestra el formulario de estado de la copia de archivos
'-----------------------------------------------------------
'
Private Sub ShowCopyDialog()
    CenterForm frmCopy
    frmCopy.Show
    frmCopy.Refresh
    If gfNoUserInput = True Then
        frmCopy.cmdExit.Visible = False
    Else
        frmCopy.cmdExit.SetFocus
    End If
End Sub

'-----------------------------------------------------------
' Procedimiento: ShowMainForm
'
' Muestra el formulario azul principal de la instalación
'-----------------------------------------------------------
'
Private Sub ShowMainForm()
    Me.Caption = gstrTitle
    Me.Show
    DrawBackGround
    Me.Refresh
End Sub

'-----------------------------------------------------------
' Procedimiento: ShowStaticMessageDialog
'
' Muestra un mensaje de instalación en una 'caja' con el
' tamaño adecuado al mensaje
'
' Entrada: [strMessage] - mensaje a visualizar
'-----------------------------------------------------------
'
Private Sub ShowStaticMessageDialog(ByVal strMessage As String)
    Dim frm As Form

    Set frm = frmMessage
    frm.lblMsg.Caption = strMessage

    '
    'El alto predeterminado es dos veces el alto del icono de instalación.
    'Si el alto del texto del mensaje es mayor, el alto del formulario se
    'incrementa a la altura de la etiqueta más
    'la mitad del alto del icono
    '
    frm.ScaleHeight = frm.imgMsg.Height * 2
    If frm.lblMsg.Height > frm.ScaleHeight Then
        frm.ScaleHeight = frm.lblMsg.Height + frm.imgMsg.Height * 0.5
    End If

    '
    'Centra el icono y la etiqueta verticalmente en el formulario
    '
    frm.imgMsg.Top = frm.ScaleHeight / 2 - frm.imgMsg.Height / 2
    frm.lblMsg.Top = frm.ScaleHeight / 2 - frm.lblMsg.Height / 2

    CenterForm frm

    frm.Show
    frm.Refresh
End Sub

'-----------------------------------------------------------
' Procedimiento: ShowWelcomeForm
'
' Muestra el formulario de bienvenida de la instalación
'-----------------------------------------------------------
'
Private Sub ShowWelcomeForm()
    If Not gfNoUserInput Then
        frmWelcome.Show 1
    End If
End Sub

'-----------------------------------------------------------
' Procedimiento: GetStrings
'
' Carga los recursos de cadena en variables globales de formularios y controles
'-----------------------------------------------------------
'
Private Sub GetStrings()
    On Error GoTo GSErr
    
    gstrSETMSG = ResolveResString(resSETMSG)
    
    Exit Sub
    
GSErr:
    MsgError mstrRESOURCELOADFAIL, MB_ICONSTOP Or MB_OK, gstrNULL
    ExitSetup Me, gintRET_FATAL
End Sub

