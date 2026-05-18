VERSION 5.00
Begin VB.Form frmRemoteServerDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4545
   ClientLeft      =   3195
   ClientTop       =   2400
   ClientWidth     =   7800
   ControlBox      =   0   'False
   Icon            =   "serverdt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdCancel 
      Caption         =   "#"
      Height          =   375
      Left            =   5580
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   3930
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "#"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3540
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   3930
      Width           =   1935
   End
   Begin VB.ComboBox cboNetworkProtocol 
      Height          =   300
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3165
      Width           =   5100
   End
   Begin VB.TextBox txtNetworkAddress 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   2535
      Width           =   5100
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   225
      TabIndex        =   7
      Top             =   1395
      Width           =   7290
      Begin VB.Label lblServerName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   240
         Width           =   7020
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblNetworkProtocol 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   3165
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNetworkAddress 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   2535
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRemoteServerDetails 
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   7020
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRemoteServerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private m_fNetworkAddressSpecified As Boolean
Private m_fNetworkProtocolSpecified As Boolean
Private m_fDCOM As Boolean

Private Declare Function RpcNetworkIsProtseqValid Lib "rpcrt4.dll" Alias "RpcNetworkIsProtseqValidA" (ByVal strProtseq As String) As Long

' Determina si una secuencia de protocolo determinada es compatible y está disponible en este equipo
Function fIsProtocolSeqSupported(ByVal strProto As String, ByVal strProtoFriendlyName) As Boolean
    Const RPC_S_OK = 0&
    Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703&
    Const RPC_S_INVALID_RPC_PROTSEQ = 1704&

    Dim rcps As Long
    Static fUnexpectedErr As Boolean

    On Error Resume Next

    fIsProtocolSeqSupported = False
    rcps = RpcNetworkIsProtseqValid(strProto)

    Select Case rcps
        Case RPC_S_OK
            fIsProtocolSeqSupported = True
        Case RPC_S_PROTSEQ_NOT_SUPPORTED
            LogNote ResolveResString(resNOTEPROTOSEQNOTSUPPORTED, "|1", strProto, "|2", strProtoFriendlyName)
        Case RPC_S_INVALID_RPC_PROTSEQ
            LogWarning ResolveResString(resNOTEPROTOSEQINVALID, "|1", strProto, "|2", strProtoFriendlyName)
        Case Else
            If Not fUnexpectedErr Then
                MsgWarning ResolveResString(resPROTOSEQUNEXPECTEDERR), vbOKOnly Or vbInformation, gstrTitle
                If gfNoUserInput Then
                    '
                    ' Probablemente es redundante ya que este formulario ya que
                    ' este formulario nunca se verá si lo ejecutamos en silencio o en modo SMS.
                    '
                    ExitSetup frmRemoteServerDetails, gintRET_FATAL
                End If
                fUnexpectedErr = True
            End If
        'End Case
    End Select
End Function

Private Sub cboNetworkProtocol_Click()
    cmdOK.Enabled = fValid()
End Sub

Private Sub cmdCancel_Click()
    ExitSetup frmRemoteServerDetails, gintRET_EXIT
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub

Private Sub Form_Load()
    Dim fMoveControlsUp As Boolean 'Indica si se mueven los controles para completar un espacio en blanco
    Dim yTopCutoff As Integer 'Todos los controles se moverán más abajo que este valor y
    
    SetFormFont Me
    Caption = ResolveResString(resREMOTESERVERDETAILSTITLE)
    lblRemoteServerDetails.Caption = ResolveResString(resREMOTESERVERDETAILSLBL)
    lblNetworkAddress.Caption = ResolveResString(resNETWORKADDRESS)
    lblNetworkProtocol.Caption = ResolveResString(resNETWORKPROTOCOL)
    cmdOK.Caption = ResolveResString(resOK)
    cmdCancel.Caption = ResolveResString(resCANCEL)
    '
    ' No nos preocupamos de los protocolos si es DCOM.
    '
    If Not m_fDCOM Then
        FillInProtocols
    End If
    
    'Activamos o desactivamos de forma selectiva los controles disponibles
    'según la información que necesitemos del usuario.
    If m_fNetworkAddressSpecified Then
        'La dirección de red ya está completada, podemos ocultar este control
        'y mover todos los controles arriba
        txtNetworkAddress.Visible = False
        lblNetworkAddress.Visible = False
        fMoveControlsUp = True
        yTopCutoff = txtNetworkAddress.Top
    ElseIf m_fNetworkProtocolSpecified Or m_fDCOM Then
        'El protocolo de red ya está completado, podemos ocultar este control
        'y mover todos los controles arriba
        cboNetworkProtocol.Visible = False
        lblNetworkProtocol.Visible = False
        fMoveControlsUp = True
        yTopCutoff = cboNetworkProtocol.Top
    End If
    
    If fMoveControlsUp Then
        'Busca cuánto hay que mover los controles hacia arriba
        Dim yDiff As Integer
        yDiff = cboNetworkProtocol.Top - txtNetworkAddress.Top
        
        Dim c As Control
        For Each c In Controls
            If c.Top > yTopCutoff Then
                c.Top = c.Top - yDiff
            End If
        Next c
        
        'Finalmente, ajusta el formulario
        Height = Height - yDiff
    End If
    
    'Centra el formulario
    Top = (Screen.Height - Height) \ 2
    Left = (Screen.Width - Width) \ 2
End Sub

'-----------------------------------------------------------
' Procedimiento: GetServerDetails
'
' Solicita del usuario cualquier información que falte de un
' servidor remoto.
'
' Entrada:
'   [strRegFile] - el nombre del archivo de registro remoto
'   [strNetworkAddress] - la dirección de red, si se conoce
'   [strNetworkProtocol] - el protocolo de red, si se conoce
'   [fDCOM] - si es true, se tiene acceso a este componente mediante
'             COM distribuido y no con Automatización remota. En este
'             caso, no es necesario el protocolo de red
'             o el nivel de autentificación.
'
' Salida:
'   [strNetworkAddress]  - la dirección de red que se pasa o que
'                          se obtiene del usuario
'   [strNetworkProtocol] - el protocolo de red que se pasa o que
'                          se obtiene del usuario
'-----------------------------------------------------------
'
Public Sub GetServerDetails( _
    ByVal strRegFile As String, _
    strNetworkAddress As String, _
    strNetworkProtocol As String, _
    fDCOM As Boolean _
)
    Dim i As Integer
    Dim strServerName As String
    
    'Comprueba que falta algo
    m_fNetworkAddressSpecified = (strNetworkAddress <> "")
    m_fNetworkProtocolSpecified = (strNetworkProtocol <> "")
    m_fDCOM = fDCOM
    
    If m_fNetworkAddressSpecified And (m_fNetworkProtocolSpecified Or m_fDCOM) Then
        'Tanto la dirección como la secuencia de protocolo de red ya se
        'han especificado en SETUP.LST. No hay que preguntar al usuario
        'más información.
        
        'A pesar de todo, es necesario comprobar que la secuencia de protocolo
        'especificada en SETUP.LST está instalada y disponible actualmente en este equipo
        '(sólo con Automatización remota).
        '
        If Not m_fDCOM Then
            CheckSpecifiedProtocolSequence strNetworkProtocol, strGetServerName(strRegFile)
        End If
        
        Exit Sub
    End If
    
    strServerName = strGetServerName(strRegFile)
    Load Me
    lblServerName.Caption = strServerName
    
    If Not gfNoUserInput Then
        '
        ' Muestra el formulario y obtiene del usuario la información necesaria
        '
        Show vbModal
    Else
        '
        ' Como es silencioso, simplemente se acepta el primero de la lista.
        '
        '
        ' Tenga en cuenta que aquí ya conocemos al menos un protocolo
        ' de la lista, si no es así el programa fallaría en el código de
        ' Form_Load al llamar a FillInProtocols().
        '
        cboNetworkProtocol.ListIndex = 0
    End If
    
    If m_fNetworkProtocolSpecified And Not m_fDCOM Then
        'La secuencia de protocolo de red ya se ha especificado en
        'SETUP.LST. Es necesario comprobar que la secuencia de protocolo especificada
        'en SETUP.LST está instalada y está disponible actualmente en esta máquina
        '(sólo 32 bits).
        CheckSpecifiedProtocolSequence strNetworkProtocol, strGetServerName(strRegFile)
    End If
    
    If Not m_fNetworkAddressSpecified Then
        strNetworkAddress = txtNetworkAddress
    End If
    If Not m_fNetworkProtocolSpecified And Not m_fDCOM Then
        strNetworkProtocol = gProtocol(cboNetworkProtocol.ListIndex + 1).strName
    End If
    
    Unload Me
End Sub

'-----------------------------------------------------------
' Procedimiento: FillInProtocols
'
' Rellena el cuadro combinado de protocolos con los protocolos
' disponibles en setup.lst
'-----------------------------------------------------------
Private Sub FillInProtocols()
    Dim i As Integer
    Dim fSuccessReading As Boolean
    
    cboNetworkProtocol.Clear
    fSuccessReading = ReadProtocols(gstrSetupInfoFile, gstrINI_SETUP)
    If Not fSuccessReading Or gcProtocols <= 0 Then
        MsgError ResolveResString(resNOPROTOCOLSINSETUPLST), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup frmRemoteServerDetails, gintRET_FATAL
    End If
    
    For i = 1 To gcProtocols
        If fIsProtocolSeqSupported(gProtocol(i).strName, gProtocol(i).strFriendlyName) Then
            cboNetworkProtocol.AddItem gProtocol(i).strFriendlyName
        End If
    Next i

    If cboNetworkProtocol.ListCount > 0 Then
        'Al menos hemos encontrado un protocolo disponible en este equipo
        Exit Sub
    End If
    
    'Ninguno de los protocolos especificados en SETUP.LST están disponibles en este equipo.
    'Es necesario informar al usuario, incluyendo el protocolo que espera.
    MsgError ResolveResString(resNOPROTOCOLSSUPPORTED1), vbExclamation Or vbOKOnly, gstrTitle
    '
    ' No registra el resto de esto en SMS. Se acepta el modo silencioso
    ' que puede tener más de 255 caracteres.
    '
    If Not gfSMS Then
        Dim strMsg As String
        strMsg = ResolveResString(resNOPROTOCOLSSUPPORTED2) & LF$
    
        For i = 1 To gcProtocols
            strMsg = strMsg & LF$ & Chr$(9) & gProtocol(i).strFriendlyName
        Next i
        
        MsgError strMsg, vbExclamation Or vbOKOnly, gstrTitle
    End If
    ExitSetup frmRemoteServerDetails, gintRET_FATAL
End Sub

'-----------------------------------------------------------
' Procedimiento: strGetServerName
'
' Si se proporciona un archivo de registro servidor remoto,
' obtiene el nombre amistoso del servidor
'-----------------------------------------------------------
Private Function strGetServerName(ByVal strRegFilename As String) As String
    Const strKey = "AppDescription="
    Dim strLine As String
    Dim iFile As Integer
    
    On Error GoTo DoErr
    
    'Esto se hace si no es posible obtener el nombre amistoso
    strGetServerName = GetFileName(strRegFilename)
    
    iFile = FreeFile
    Open strRegFilename For Input Access Read Lock Read Write As #iFile
    While Not EOF(iFile)
        Line Input #iFile, strLine
        If Left$(strLine, Len(strKey)) = strKey Then
            'Hemos encontrado la línea con el nombre amistoso del servidor
            Dim strName As String
            strName = Mid$(strLine, Len(strKey) + 1)
            If strName <> "" Then
                strGetServerName = strName
            End If
            Close iFile
            Exit Function
        End If
    Wend
    Close iFile
    
    Exit Function
DoErr:
    strGetServerName = ""
End Function

Private Sub txtNetworkAddress_Change()
    cmdOK.Enabled = fValid()
End Sub

'Devuelve True si la entrada es válida
Private Function fValid() As Boolean
    fValid = True
    '
    ' Si esto es dcom, no nos preocupamos del protocolo de red.
    '
    If m_fDCOM = False Then
        If Not m_fNetworkProtocolSpecified And (cboNetworkProtocol.ListIndex < 0) Then
            fValid = False
        End If
    End If
    
    If Not m_fNetworkAddressSpecified And (txtNetworkAddress = "") Then
        fValid = False
    End If
End Function

Private Sub CheckSpecifiedProtocolSequence(ByVal strNetworkProtocol As String, ByVal strFriendlyServerName As String)
        'Se intenta encontrar el nombre amistoso de este protocolo en la lista de SETUP.LST
        Dim fSuccessReading As Boolean
        Dim strFriendlyName As String
        Dim i As Integer
        
        strFriendlyName = strNetworkProtocol 'Hay que hacer esto si no se encuentra nada mejor
        
        fSuccessReading = ReadProtocols(gstrSetupInfoFile, gstrINI_SETUP)
        If fSuccessReading And gcProtocols > 0 Then
            For i = 1 To gcProtocols
                If gProtocol(i).strName = strNetworkProtocol Then
                    strFriendlyName = gProtocol(i).strFriendlyName
                    Exit For
                End If
            Next i
        End If
        
        'Comprueba si el protocolo está disponible
        If fIsProtocolSeqSupported(strNetworkProtocol, strFriendlyName) Then
            'Está disponible
            Exit Sub
        Else
            'No está disponible. Se informa con un mensaje y se continúa con la instalación.
Retry:
            If gfNoUserInput Or MsgError( _
                ResolveResString(resSELECTEDPROTONOTSUPPORTED, "|1", strFriendlyServerName, "|2", strFriendlyName), _
                vbInformation Or vbOKCancel, _
                gstrTitle) _
              = vbCancel Then
                '
                ' El usuario eligió Cancelar. Hay que darle una oportunidad de salir (si no está en modo silencioso o SMS);
                ' de otra forma cualquier llamada a ExitSetup es fatal.
                '
                ExitSetup frmRemoteServerDetails, gintRET_EXIT
                GoTo Retry
            Else
                'El usuario eligió Aceptar. Continúa la instalación.
                Exit Sub
            End If
        End If
End Sub

