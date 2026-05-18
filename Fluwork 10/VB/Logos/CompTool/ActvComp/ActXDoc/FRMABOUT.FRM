VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de PrimeDoc"
   ClientHeight    =   3555
   ClientLeft      =   1320
   ClientTop       =   2100
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3525
      TabIndex        =   0
      Top             =   2625
      Width           =   1980
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "I&nformación del sistema..."
      Height          =   345
      Left            =   3540
      TabIndex        =   2
      Top             =   3075
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "Documento ActiveX"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Documento ActiveX FirstDoc"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de claves del registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de claves del registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                     ' Cadena Unicode terminada en nulo
Const REG_DWORD = 4                  ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "Acerca de " & App.Title
  lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
  lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
  On Error GoTo SysInfoErr

  Dim rc As Long
  Dim SysInfoPath As String
  
  ' Intenta obtener del Registro la ruta\nombre de programa en la
  ' información del sistema...
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
  ' Intenta obtener del Registro sólo la ruta de programa en la
  ' información del sistema...
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
      ' Valida la existencia de una versión de archivo de 32 bits conocida
      If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
          SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
          
      ' Error - No se encuentra el archivo...
      Else
          GoTo SysInfoErr
      End If
  ' Error - No se encuentra la entrada del Registro...
  Else
      GoTo SysInfoErr
  End If
  
  Call Shell(SysInfoPath, vbNormalFocus)
  
  Exit Sub
SysInfoErr:
  MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
  Dim i As Long                               ' Contador del bucle
  Dim rc As Long                              ' Código de retorno
  Dim hKey As Long                            ' Controlador de una clave del registro abierta
  Dim hDepth As Long                          '
  Dim KeyValType As Long                      ' Tipo de datos de una clave del Registro
  Dim tmpVal As String                        ' Almacenamiento temporal de un valor de clave del Registro
  Dim KeyValSize As Long                      ' Tamaño de una variable de clave del Registro
  '------------------------------------------------------------
  ' Abre RegKey bajo KeyRoot {HKEY_LOCAL_MACHINE...}
  '------------------------------------------------------------
  rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abre una clave del Registro
  
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError     ' Trata el error...
  
  tmpVal = String$(1024, 0)                          ' Asigna espacio de variable
  KeyValSize = 1024                                  ' Marca tamaño variable
  
  '------------------------------------------------------------
  ' Recupera un valor de clave del Registro...
  '------------------------------------------------------------
  rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                       KeyValType, tmpVal, KeyValSize)  ' Obtiene/crea valor de clave
                      
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError        ' Trata los errores
  
  If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then         ' Win95 agrega cadena terminada en Null...
      tmpVal = Left(tmpVal, KeyValSize - 1)             ' Encontrado Null, extraer de la cadena
  Else                                                  ' WinNT NO termina la cadena con Null...
      tmpVal = Left(tmpVal, KeyValSize)                 ' No encontrado Null, extraer sólo la cadena
  End If
  '--------------------------------------------------------------
  ' Determina el tipo del valor de la clave para su conversión...
  '--------------------------------------------------------------
  Select Case KeyValType                                ' Busca los tipos de datos...
  Case REG_SZ                                           ' Tipo de datos de la clave del Registro
      KeyVal = tmpVal                                   ' Copia el valor de cadena
  Case REG_DWORD                                        ' Tipo de datos de clave de Registro Double Word
      For i = Len(tmpVal) To 1 Step -1                  ' Convierte cada bit
          KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1))) ' Genera valor Char. a Char.
      Next
      KeyVal = Format$("&h" + KeyVal)                   ' Convierte Double Word en String
  End Select
  
  GetKeyValue = True                                    ' Devuelve éxito
  rc = RegCloseKey(hKey)                                ' Cierra la clave del Registro
  Exit Function                                         ' Sale
  
GetKeyError:    ' Terminación cuando se ha producido un error...
  KeyVal = ""                                           ' Cadena vacía como valor devuelto
  GetKeyValue = False                                   ' Devuelve fallo
  rc = RegCloseKey(hKey)                                ' Cierra la clave del Registro
End Function
