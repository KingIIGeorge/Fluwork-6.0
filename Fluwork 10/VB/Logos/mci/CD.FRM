VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmCD 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   225
   ClientTop       =   2700
   ClientWidth     =   8415
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8415
   Begin VB.PictureBox picCD4 
      BackColor       =   &H00808080&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   -15
      Width           =   1570
      Begin VB.PictureBox picCD5 
         BackColor       =   &H00575757&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   1245
         TabIndex        =   12
         Top             =   110
         Width           =   1300
      End
   End
   Begin VB.PictureBox picCD6 
      BackColor       =   &H00808080&
      Height          =   3255
      Left            =   6845
      ScaleHeight     =   3195
      ScaleWidth      =   1515
      TabIndex        =   13
      Top             =   -15
      Width           =   1575
      Begin VB.PictureBox picCD7 
         BackColor       =   &H00575757&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   1245
         TabIndex        =   14
         Top             =   110
         Width           =   1300
      End
   End
   Begin VB.Frame fraCD1 
      BackColor       =   &H00808080&
      Caption         =   "Reproductor de CD con MCI"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.PictureBox picCD3 
         Height          =   255
         Left            =   720
         Picture         =   "cd.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraCD4 
         BackColor       =   &H00808080&
         Height          =   975
         Left            =   2160
         TabIndex        =   5
         Top             =   2040
         Width           =   2895
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Cargar"
            Height          =   255
            Left            =   1920
            TabIndex        =   6
            Top             =   600
            Width           =   855
         End
         Begin VB.PictureBox picCD1 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            Picture         =   "cd.frx":09DA
            ScaleHeight     =   255
            ScaleWidth      =   2655
            TabIndex        =   7
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblIndicator 
            BackColor       =   &H00404040&
            Height          =   135
            Left            =   1485
            TabIndex        =   10
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.PictureBox picCD2 
         Height          =   255
         Left            =   720
         Picture         =   "cd.frx":1078
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraCD2 
         BackColor       =   &H00808080&
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4815
         Begin VB.Frame fraCD3 
            BackColor       =   &H00808080&
            Height          =   855
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   855
            Begin VB.Label lblTrackCaption 
               BackColor       =   &H00808080&
               Caption         =   " Pista   "
               Height          =   255
               Left            =   180
               TabIndex        =   4
               Top             =   480
               Width           =   495
            End
            Begin VB.Label lblTrack 
               Alignment       =   2  'Center
               Caption         =   "0"
               Height          =   255
               Left            =   240
               TabIndex        =   3
               Top             =   240
               Width           =   375
            End
         End
         Begin MCI.MMControl mciCDPlayer 
            Height          =   735
            Left            =   1200
            TabIndex        =   15
            Top             =   360
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   1296
            _Version        =   327680
            DeviceType      =   ""
            FileName        =   ""
         End
      End
   End
End
Attribute VB_Name = "frmCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLoad_Click()
    ' Abre el dispositivo CD. El disco debe estar ya en la unidad.
    On Error GoTo MCI_ERROR
    mciCDPlayer.Command = "Open"
    On Error GoTo 0
    
    ' Establece el formato del tiempo.
    mciCDPlayer.TimeFormat = vbMCIFormatTmsf
    
    ' Deshabilita el botón de comando Cargar, y muestra el mapa de bits
    ' "disco en la unidad".
    cmdLoad.Enabled = False
    picCD1.Picture = picCD3.Picture
    
    ' Establece la primera pista en el número de pista.
    lblTrack.Caption = "1"
    Exit Sub

MCI_ERROR:
    DisplayErrorMessageBox
    Resume MCI_EXIT

MCI_EXIT:
    Unload frmCD
End Sub

Private Sub Form_Load()
    mciCDPlayer.Wait = True
    
    mciCDPlayer.UpdateInterval = 0
    
    ' Establece un dispositivo CD musical en la propiedad
    ' DeviceType.
    mciCDPlayer.DeviceType = "CDAudio"

    ' Establece 0 (predeterminado) en el número de pista.
    lblTrack.Caption = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMCITest.Show
End Sub

Private Sub mciCDPlayer_EjectClick(Cancel As Integer)
    ' Habilita el botón de comando Cargar, y muestra el mapa de bits
    ' "disco fuera de la unidad".
    cmdLoad.Enabled = True
    picCD1.Picture = picCD2.Picture
    mciCDPlayer.UpdateInterval = 0
    
    ' Expulsa el disco de la unidad CD, y la cierra.
    On Error GoTo MCI_ERROR2
    mciCDPlayer.Command = "Eject"
    mciCDPlayer.Command = "Close"
    On Error GoTo 0
    
    ' Restablece 0 como número de pista
    lblTrack.Caption = "0"
    
    ' Desactiva el "indicador de reproducción".
    lblIndicator.BackColor = &H404040
    Exit Sub

MCI_ERROR2:
    DisplayErrorMessageBox
    Resume Next
End Sub

Private Sub mciCDPlayer_NextCompleted(ErrorCode As Long)
    ' Establece la nueva pista en el número de pista.
    lblTrack.Caption = Str$(mciCDPlayer.Track)
End Sub

Private Sub mciCDPlayer_PauseClick(Cancel As Integer)
    mciCDPlayer.UpdateInterval = 0
        
    ' Desactiva el "indicador de reproducción".
    lblIndicator.BackColor = &H404040
End Sub

Private Sub mciCDPlayer_PlayClick(Cancel As Integer)
    mciCDPlayer.UpdateInterval = 1000
    
    ' Establece la pista actual en el número de pista.
    lblTrack.Caption = Str$(mciCDPlayer.Track)
End Sub

Private Sub mciCDPlayer_PrevCompleted(ErrorCode As Long)
    ' Establece la nueva pista en el número de pista.
    lblTrack.Caption = Str$(mciCDPlayer.Track)
End Sub

Private Sub mciCDPlayer_StatusUpdate()
    ' Establece la pista actual en el número de pista.
    lblTrack.Caption = Str$(mciCDPlayer.Track)

    ' Determina si el "indicador de reproducción" está
    ' activado o desactivado, dependiendo de si el dispositivo
    ' está reproduciendo actualmente.
    If mciCDPlayer.Mode = vbMCIModePlay Then
        lblIndicator.BackColor = &H80FF&
    Else
        lblIndicator.BackColor = &H404040
    End If
End Sub

Private Sub mciCDPlayer_StopClick(Cancel As Integer)
    mciCDPlayer.UpdateInterval = 0
    
    ' Desactiva el "indicador de reproducción".
    lblIndicator.BackColor = &H404040

    ' Restablece la pista 1 del CD.
    mciCDPlayer.To = mciCDPlayer.Start
    mciCDPlayer.Command = "Seek"
    mciCDPlayer.Track = 1
    lblTrack.Caption = "1"
End Sub

