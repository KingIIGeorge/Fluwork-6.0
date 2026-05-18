VERSION 5.00
Begin VB.Form frmGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "* #"
   ClientHeight    =   5250
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5460
   Icon            =   "group.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   Begin VB.DirListBox dir95Groups 
      Height          =   930
      Left            =   765
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   105
      TabIndex        =   7
      Top             =   4650
      Width           =   5220
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "#"
      Height          =   345
      Left            =   2880
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   4800
      Width           =   1140
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "#"
      Default         =   -1  'True
      Height          =   345
      Left            =   1395
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   4800
      Width           =   1140
   End
   Begin VB.ListBox lstGroups 
      Height          =   2040
      ItemData        =   "group.frx":0442
      Left            =   1080
      List            =   "group.frx":0449
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2220
      Width           =   3240
   End
   Begin VB.TextBox txtGroup 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Text            =   "*"
      Top             =   1410
      Width           =   3270
   End
   Begin VB.Label lblDDE 
      Height          =   225
      Left            =   225
      TabIndex        =   9
      Top             =   1350
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblGroups 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1950
      Width           =   105
   End
   Begin VB.Label lblGroup 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1170
      Width           =   105
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   165
      Width           =   5100
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrGroup As String
Dim mstrDefGroup As String
Dim mstrProgramsPath As String
Dim mfrm As Form

Private Sub cmdCancel_Click()
    ExitSetup frmGroup, gintRET_EXIT
End Sub

Private Sub cmdContinue_Click()
    mstrGroup = txtGroup.Text
    If Not fCreateProgGroup() Then
        '
        ' Imposible crear el grupo. Hay que
        ' dejar que lo intente el usuario.
        '
        txtGroup.SetFocus
    Else
        '
        ' El grupo se creó correctamente, se descarga el cuadro de diálogo
        ' Elegir grupo de programas y se continúa con la instalación.
        '
        Unload Me
    End If
End Sub
Private Function fCreateProgGroup() As Boolean
'
' Crea un grupo de programas para NT o Win95.
'
    Dim strMsg As String
    If TreatAsWin95() Then
        If Not fValidFilename(mstrGroup) Then
            strMsg = ResolveResString(resGROUPINVALIDGROUPNAME, "|1", CStr(gintMAX_PATH_LEN), "|2", ResolveResString(resCOMMON_INVALIDFILECHARS))
            MsgFunc strMsg, vbOKOnly Or vbQuestion, gstrTitle
            GoTo CGError
        End If
    Else
        If Not fValidNTGroupName(mstrGroup) Then
            strMsg = ResolveResString(resGROUPINVALIDGROUPNAME, "|1", CStr(gintMAX_GROUPNAME_LEN%), "|2", ResolveResString(resGROUPINVALIDCHARS))
            MsgFunc strMsg, vbOKOnly Or vbQuestion, gstrTitle
            GoTo CGError
        End If
            
    End If
    '
    'Continúa y crea el grupo de programas principal
    '
    If Not fCreateOSProgramGroup(mfrm, mstrGroup, True) Then
        GoTo CGError
    End If
    
    fCreateProgGroup = True
    Exit Function
CGError:
    fCreateProgGroup = False
End Function
Private Sub Form_Load()
    '
    ' Inicializa las propiedades del control localizado.
    '
    SetFormFont Me
    Me.Caption = ResolveResString(resGROUPFRM, "|1", gstrAppName)
    lblMain.Caption = ResolveResString(resGROUPLBLMAIN)
    lblGroup.Caption = ResolveResString(resGROUPLBLGROUP)
    lblGroups.Caption = ResolveResString(resGROUPLBLGROUPS)
    cmdContinue.Caption = ResolveResString(resGROUPBTNCONTINUE)
    cmdCancel.Caption = ResolveResString(resLOG_IDCANCEL)
    '
    ' Inicializa el cuadro de texto de Grupo de programa con
    ' el título de la aplicación.
    '
    txtGroup.Text = gstrTitle
    '
    ' Carga el control ListBox con los grupos del Administrador de programas.
    '
    If TreatAsWin95() Then
        LoadW95Groups
    Else
        LoadProgManGroups
    End If
    '
    ' Inicializa el cuadro de texto de Grupo de programas con el
    ' grupo predeterminado seleccionado en el cuadro de lista.
    '
    txtGroup.Text = lstGroups.List(lstGroups.ListIndex)
End Sub

Private Sub lstGroups_Click()
    txtGroup.Text = lstGroups.List(lstGroups.ListIndex)
End Sub

Private Sub txtGroup_Change()
    cmdContinue.Enabled = Len(Trim(txtGroup.Text)) > 0
End Sub
Sub LoadProgManGroups()
'
' Esta rutina usa DDE para hablar con el Administrador de programas
' y recuperar una lista de todos los grupos que administra.
' Se debe llamar si el shell es NT 3.51.
' Si es Win95 o NT4, llama en su lugar a
' LoadW95Groups().
'
' Cadenas especiales que usa esta rutina.
' No hay que traducir estas cadenas.
'
    Const strPROGMANLINKTOPIC = "ProgMan|Progman"
    Const strPROGMANLINKITEM = "Progman"
    Const strNDWGROUP = "Quick Access"

    Dim strGroups As String
    Dim strGroup As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim iGroup As Long
    
    lblDDE.LinkTopic = strPROGMANLINKTOPIC
    lblDDE.LinkItem = strPROGMANLINKITEM
    lblDDE.LinkMode = 2
    lblDDE.LinkRequest
    On Error Resume Next
    lblDDE.LinkMode = 0
    '
    ' La llamada DDE sólo pone los nombres de todos los grupos
    ' en la propiedad caption del control lblDDE.
    ' Queremos transferirlos al cuadro de lista. Están separados
    ' por retornos de carro.
    '
    strGroups = lblDDE.Caption
    intAnchor = 1
    intOffset = InStr(intAnchor, strGroups, CRLF)
    lstGroups.Clear
    Do While intOffset > 0
        strGroup = Mid(strGroups, intAnchor, intOffset - intAnchor)
        '
        ' Norton Desktop for Windows usa el grupo "Quick Access" que
        ' sustituye al Administrador de programas. Si intenta agregar iconos
        ' a este grupo, fallará cuando realizemos una llamada DDE.linkrequest.
        ' Por lo tanto, se salta este grupo.
        '
        If strGroup <> strNDWGROUP Then
            lstGroups.AddItem strGroup
        End If
        
        intAnchor = intOffset + 2
        intOffset = InStr(intAnchor, strGroups, CRLF)
    Loop
    '
    ' El cuadro de lista lstGroups contiene ahora una lista de todos
    ' los grupos del administrador de programas.
    '
    ' Busca el grupo predeterminado en la lista y los selecciona.
    ' Si no está, lo agrega.
    '
    iGroup = SendMessageString(lstGroups.hwnd, LB_FINDSTRINGEXACT, -1, mstrDefGroup)
    If iGroup = LB_ERR Then
        '
        ' El grupo todavía no existe, se agrega a la lista.
        '
        lstGroups.AddItem mstrDefGroup
        lstGroups.ListIndex = lstGroups.NewIndex
    Else
        lstGroups.ListIndex = iGroup
    End If
End Sub
Sub LoadW95Groups()
'
' Esta rutina usa el registro del sistema para recuperar
' una lista de todas las subcarpetas que están en la carpeta
' \windows\start menu\programs.
' Sólo se puede llamar si el shell es Win95 o
' NT4. Si es NT 3.51, llama en su lugar a
' LoadProgManGroups().
'
    Dim strFolder As String
    Dim iFolder As Integer
    
    mstrProgramsPath = strGetProgramsFilesPath()
    strFolder = Dir(mstrProgramsPath, vbDirectory)   ' Recupera la primera entrada.
    lstGroups.Clear
    Do While strFolder <> ""
        '
        ' Ignora el directorio actual y su directorio.
        '
        If strFolder <> "." And strFolder <> ".." Then
            '
            ' Comprueba que tenemos un directorio y no un archivo.
            '
            If (GetAttr(mstrProgramsPath & strFolder) And vbDirectory) = vbDirectory Then
                '
                ' Tenemos un directorio, se agrega a la lista.
                '
                lstGroups.AddItem strFolder
            End If
        End If
        '
        ' Obtiene la siguiente subcarpeta en la carpeta de Programas
        '
        strFolder = Dir
    Loop
    '
    ' EL cuadro de lista lstGroups contiene ahora una lista de todas las
    ' subcarpetas de programas (los grupos).
    '
    ' Busca la carpeta predeterminada de la lista y la selecciona.
    ' Si no está, la agrega.
    '
    iFolder = SendMessageString(lstGroups.hwnd, LB_FINDSTRINGEXACT, -1, mstrDefGroup)
    If iFolder = LB_ERR Then
        '
        ' El grupo ya no existe, se agrega a la lista.
        '
        lstGroups.AddItem mstrDefGroup
        lstGroups.ListIndex = lstGroups.NewIndex
    Else
        lstGroups.ListIndex = iFolder
    End If
End Sub
Public Property Get GroupName(frm As Form, strDefGroup As String) As String
    mstrDefGroup = strDefGroup
    Set mfrm = frm
    
    If gfNoUserInput = True Then
        mstrGroup = mstrDefGroup
        If Not fCreateProgGroup() Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    Else
        Me.Show vbModal
    End If
    GroupName = mstrGroup
End Property

Private Sub txtGroup_GotFocus()
    txtGroup.SelStart = 0
    txtGroup.SelLength = 32767
End Sub
