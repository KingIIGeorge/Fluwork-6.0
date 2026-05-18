VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propiedades de CommPort"
   ClientHeight    =   4260
   ClientLeft      =   4140
   ClientTop       =   1665
   ClientWidth     =   6135
   Icon            =   "frmProps.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   570
      Width           =   5565
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   4455
         TabIndex        =   22
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Frame Frame1 
         Caption         =   "Velocidad máxima"
         Height          =   870
         Left            =   180
         TabIndex        =   20
         Top             =   630
         Width           =   2565
         Begin VB.ComboBox cboSpeed 
            Height          =   315
            Left            =   375
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame fraConnection 
         Caption         =   "Preferencias de la conexión"
         Height          =   1770
         Left            =   180
         TabIndex        =   12
         Top             =   1635
         Width           =   2565
         Begin VB.ComboBox cboStopBits 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1260
            Width           =   1140
         End
         Begin VB.ComboBox cboParity 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   810
            Width           =   1140
         End
         Begin VB.ComboBox cboDataBits 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   330
            Width           =   1140
         End
         Begin VB.Label Label5 
            Caption         =   "Bits de parada:"
            Height          =   285
            Left            =   180
            TabIndex        =   19
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label4 
            Caption         =   "Paridad:"
            Height          =   285
            Left            =   180
            TabIndex        =   18
            Top             =   855
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Bits de datos:"
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   375
            Width           =   1170
         End
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   300
         Left            =   4455
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   705
         Width           =   1080
      End
      Begin VB.Frame Frame7 
         Caption         =   "&Eco"
         Height          =   870
         Left            =   2835
         TabIndex        =   7
         Top             =   630
         Width           =   1530
         Begin VB.OptionButton optEcho 
            Caption         =   "Desactivado"
            Height          =   255
            Index           =   0
            Left            =   120
            MaskColor       =   &H00000000&
            TabIndex        =   9
            Top             =   300
            Width           =   1320
         End
         Begin VB.OptionButton optEcho 
            Caption         =   "Activado"
            Height          =   195
            Index           =   1
            Left            =   135
            MaskColor       =   &H00000000&
            TabIndex        =   8
            Top             =   585
            Width           =   1155
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Control de &flujo"
         Height          =   1770
         Left            =   2835
         TabIndex        =   2
         Top             =   1635
         Width           =   1530
         Begin VB.OptionButton optFlow 
            Caption         =   "Ninguno"
            Height          =   255
            Index           =   0
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   6
            Top             =   345
            Width           =   1035
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "Xon/Xoff"
            Height          =   255
            Index           =   1
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   5
            Top             =   645
            Width           =   1035
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "RTS"
            Height          =   255
            Index           =   2
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   4
            Top             =   945
            Width           =   735
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "Xon/RTS"
            Height          =   255
            Index           =   3
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   3
            Top             =   1245
            Width           =   1050
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Puerto:"
         Height          =   315
         Left            =   330
         TabIndex        =   13
         Top             =   180
         Width           =   495
      End
   End
   Begin ComctlLib.TabStrip tabSettings 
      Height          =   4065
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   7170
      _Version        =   327680
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Propiedades"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iFlow As Integer, iTempEcho As Boolean


Sub LoadPropertySettings()
Dim i As Integer, Settings As String, Offset As Integer

' Carga las configuraciones de puerto
For i = 1 To 16
    cboPort.AddItem "Com" & Trim$(Str$(i))
Next i

' Carga las configuraciones de velocidad
cboSpeed.AddItem "110"
cboSpeed.AddItem "300"
cboSpeed.AddItem "600"
cboSpeed.AddItem "1200"
cboSpeed.AddItem "2400"
cboSpeed.AddItem "4800"
cboSpeed.AddItem "9600"
cboSpeed.AddItem "14400"
cboSpeed.AddItem "19200"
cboSpeed.AddItem "28800"
cboSpeed.AddItem "38400"
cboSpeed.AddItem "56000"
cboSpeed.AddItem "57600"
cboSpeed.AddItem "115200"
cboSpeed.AddItem "128000"
cboSpeed.AddItem "256000"

' Carga las configuraciones de bits de datos
cboDataBits.AddItem "4"
cboDataBits.AddItem "5"
cboDataBits.AddItem "6"
cboDataBits.AddItem "7"
cboDataBits.AddItem "8"

' Carga las configuraciones de paridad
cboParity.AddItem "Even"
cboParity.AddItem "Odd"
cboParity.AddItem "None"
cboParity.AddItem "Mark"
cboParity.AddItem "Space"

' Carga las configuraciones de bits de parada
cboStopBits.AddItem "1"
cboStopBits.AddItem "1.5"
cboStopBits.AddItem "2"

' Establece la configuración predeterminada

Settings = frmTerminal.MSComm1.Settings

' En todos los casos, el componente más a la derecha de Settings
' será un solo carácter, excepto cuando haya 1,5 bits de parada.
If InStr(Settings, ".") > 0 Then
    Offset = 2
Else
    Offset = 0
End If

cboSpeed.Text = Left$(Settings, Len(Settings) - 6 - Offset)
Select Case Mid$(Settings, Len(Settings) - 4 - Offset, 1)
Case "e"
    cboParity.ListIndex = 0
Case "m"
    cboParity.ListIndex = 1
Case "n"
    cboParity.ListIndex = 2
Case "o"
    cboParity.ListIndex = 3
Case "s"
    cboParity.ListIndex = 4
End Select

cboDataBits.Text = Mid$(Settings, Len(Settings) - 2 - Offset, 1)
cboStopBits.Text = Right$(Settings, 1 + Offset)
    
cboPort.ListIndex = frmTerminal.MSComm1.CommPort - 1

optFlow(frmTerminal.MSComm1.Handshaking).Value = True
If Echo Then
    optEcho(1).Value = True
Else
    optEcho(0).Value = True
End If

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim OldPort As Integer, ReOpen As Boolean

On Error Resume Next

Echo = iTempEcho
OldPort = frmTerminal.MSComm1.CommPort
NewPort = cboPort.ListIndex + 1

If NewPort <> OldPort Then    ' Si cambia el número de puerto, cierra el antiguo.
    If frmTerminal.MSComm1.PortOpen Then
           frmTerminal.MSComm1.PortOpen = False
           ReOpen = True
    End If

    frmTerminal.MSComm1.CommPort = NewPort     ' Establece el nuevo número de puerto.
    
    If Err = 0 Then
        If ReOpen Then
            frmTerminal.MSComm1.PortOpen = True
            frmTerminal.mnuOpen.Checked = frmTerminal.MSComm1.PortOpen
            frmTerminal.mnuSendText.Enabled = frmTerminal.MSComm1.PortOpen
            frmTerminal.tbrToolBar.Buttons("TransmitTextFile").Enabled = frmTerminal.MSComm1.PortOpen
        End If
    End If
        
    If Err Then
        MsgBox Error$, 48
        frmTerminal.MSComm1.CommPort = OldPort
        Exit Sub
    End If
End If


frmTerminal.MSComm1.Settings = Trim$(cboSpeed.Text) & "," & Left$(cboParity.Text, 1) _
    & "," & Trim$(cboDataBits.Text) & "," & Trim$(cboStopBits.Text)

If Err Then
    MsgBox Error$, 48
    Exit Sub
End If

frmTerminal.MSComm1.Handshaking = iFlow
If Err Then
    MsgBox Error$, 48
    Exit Sub
End If

SaveSetting App.Title, "Properties", "Settings", frmTerminal.MSComm1.Settings
SaveSetting App.Title, "Properties", "CommPort", frmTerminal.MSComm1.CommPort
SaveSetting App.Title, "Properties", "Handshaking", frmTerminal.MSComm1.Handshaking
SaveSetting App.Title, "Properties", "Echo", Echo

Unload Me

End Sub

Private Sub Form_Load()

' Establece el tamaño del formulario
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

' Ajusta el tamaño del marco al control tabstrip
fraSettings.Move tabSettings.ClientLeft, tabSettings.ClientTop

' Se asegura de que el marco sea el control situado encima
fraSettings.ZOrder

' Carga la configuración actual de las propiedades
LoadPropertySettings

End Sub




Private Sub optEcho_Click(Index As Integer)
If Index = 1 Then
    iTempEcho = True
Else
    iTempEcho = False
End If
End Sub

Private Sub optFlow_Click(Index As Integer)
iFlow = Index
End Sub


