VERSION 5.00
Begin VB.Form frmCallDlls 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Llamadas a procedimientos en DLL"
   ClientHeight    =   1920
   ClientLeft      =   4305
   ClientTop       =   5280
   ClientWidth     =   5520
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   960
      Picture         =   "CALLDLLS.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   32
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picCopy 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1680
      ScaleHeight     =   33
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   33
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   240
      Picture         =   "CALLDLLS.frx":030A
      ScaleHeight     =   32
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   32
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton cmdBitBlt 
      Caption         =   "BitBlt"
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Instrucciones"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label lblInfo 
         Caption         =   "Haga clic con el botón secundario del mouse en los iconos de la derecha."
         Height          =   855
         Index           =   1
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Timer tmrBounce 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Sistema operativo"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "(Modo mejorado)"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Disk Operating System 5.0"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2265
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Microsoft Windows Versión 3.1"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Información general"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblInfo 
         Caption         =   "Teclado:"
         Height          =   435
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   4230
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Lenguaje:"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Mouse:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Red:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Procesador, memoria y recursos del sistema"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Timer tmrSysInfo 
         Interval        =   1
         Left            =   3840
         Top             =   240
      End
      Begin VB.Shape shpFrame 
         Height          =   255
         Index           =   3
         Left            =   1080
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Shape shpBar 
         BackStyle       =   1  'Opaque
         DrawMode        =   7  'Invert
         Height          =   255
         Index           =   3
         Left            =   1080
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblResInfo 
         Alignment       =   2  'Center
         Caption         =   "archivo"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   27
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblR 
         Caption         =   "Archivo"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape shpFrame 
         Height          =   255
         Index           =   1
         Left            =   1080
         Top             =   840
         Width           =   3135
      End
      Begin VB.Shape shpBar 
         BackStyle       =   1  'Opaque
         DrawMode        =   7  'Invert
         Height          =   255
         Index           =   1
         Left            =   1080
         Top             =   840
         Width           =   1695
      End
      Begin VB.Shape shpFrame 
         Height          =   255
         Index           =   2
         Left            =   1080
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Shape shpBar 
         BackStyle       =   1  'Opaque
         DrawMode        =   7  'Invert
         Height          =   255
         Index           =   2
         Left            =   1080
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblResInfo 
         Alignment       =   2  'Center
         Caption         =   "usuario"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lblResInfo 
         Alignment       =   2  'Center
         Caption         =   "física"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblR 
         Caption         =   "Física"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblR 
         Caption         =   "Virtual"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "CPU: 486 (con procesador matemático)"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Memoria libre (KB)"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Vídeo"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Colores:"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Resolución"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Caption         =   "Controlador de vídeo:"
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Image ImgIcon 
      Height          =   480
      Index           =   1
      Left            =   4800
      Picture         =   "CALLDLLS.frx":0614
      Top             =   720
      Width           =   480
   End
   Begin VB.Image ImgIcon 
      Height          =   480
      Index           =   0
      Left            =   4800
      Picture         =   "CALLDLLS.frx":091E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmCallDlls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dx As Integer, dy As Integer, X As Integer, Y As Integer
Dim PicWidth As Integer, PicHeight As Integer
Dim RightEdge As Integer, BottomEdge As Integer

Private Sub cmdBitBlt_Click()
Dim t As Integer
    If tmrBounce.Enabled Then
        tmrBounce.Enabled = False
        Refresh
    Else
        ScaleMode = PIXELS
        dx = 15
        dy = 15
        tmrBounce.Enabled = True
        PicWidth = picSprite.ScaleWidth
        PicHeight = picSprite.ScaleHeight
        picCopy.Width = PicWidth
        picCopy.Height = PicHeight
        t = BitBlt(picCopy.hDC, 0, 0, PicWidth, PicHeight, hDC, X, Y, SRCCOPY)
    End If
End Sub

Private Sub FillSysInfo()
    Dim FreeSpace As Currency, FreeBlock As Currency, strTmp As String
    Dim YourMem As MEMORYSTATUS, myVer As MYVERSION
    
    fraInfo(4).Visible = False
   
    'Información acerca del sistema operativo
        Dim YourSystem As SystemInfo
        GetSystemInfo YourSystem
        
        myVer = WindowsVersion()
        
        If myVer.lMajorVersion = 4 Then
            If myVer.lExtraInfo = VER_PLATFORM_WIN32_NT Then
                strTmp = "Windows NT versión : "
            ElseIf myVer.lExtraInfo = VER_PLATFORM_WIN32_WINDOWS Then
                strTmp = "Windows 95 versión : "
            End If
        Else
            strTmp = "Windows versión : "
        End If
            
        lblInfo(2).Caption = strTmp & myVer.lMajorVersion & "." & myVer.lMinorVersion
        lblInfo(3).Caption = ""
        lblInfo(4).Caption = ""
                
    ' Información acerca de la CPU
        lblInfo(5).Caption = "CPU: " & YourSystem.dwProcessorType
    ' Información de vídeo
        lblInfo(8).Visible = False
        lblInfo(9).Caption = "Resolución: " & Screen.Width \ Screen.TwipsPerPixelX & " x " & Screen.Height \ Screen.TwipsPerPixelY
        lblInfo(10).Caption = "Colores: " & DeviceColors((hDC))
    
        ' Información general
        If GetSystemMetrics(SM_MOUSEPRESENT) Then
            lblInfo(11).Caption = "Mouse: " & GetSysIni("boot.description", "mouse.drv")
        Else
            lblInfo(11).Caption = "Sin mouse"
        End If
            lblInfo(12).Caption = "Red: " & GetSysIni("boot.description", "network.drv")
            lblInfo(13).Caption = "Lenguaje: " & GetSysIni("boot.description", "language.dll")
            lblInfo(14).Caption = "Teclado: " & GetSysIni("boot.description", "keyboard.typ")
End Sub

Private Sub Form_Load()
    Show                ' Asegurar que este formulario tenga un controlador, etc.
    Load frmMenus
    Icon = ImgIcon(1).Picture
    FillSysInfo
    frmMenus.mnuSysInfo(3).Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Resize()
Dim t As Integer, hDC As Long
    If WindowState = MINIMIZED Then
        RightEdge = Screen.Width \ Screen.TwipsPerPixelX
        BottomEdge = Screen.Height \ Screen.TwipsPerPixelY
        If tmrBounce.Enabled Then
            hDC = GetDC(GetDesktopWindow())
            t = BitBlt(picCopy.hDC, 0, 0, PicWidth, PicHeight, hDC, X, Y, SRCCOPY)
            ReleaseDC GetDesktopWindow(), hDC
        End If
    Else
        ScaleMode = PIXELS
        RightEdge = ScaleWidth
        BottomEdge = ScaleHeight
        If tmrBounce.Enabled Then
            hDC = GetDC(GetDesktopWindow())
            t = BitBlt(hDC, X, Y, PicWidth, PicHeight, picCopy.hDC, 0, 0, SRCCOPY)
            ReleaseDC GetDesktopWindow(), hDC
        End If

    End If
End Sub

Private Sub ImgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IX  As Integer, IY As Integer
'hMenu y hSubMenu deben ser Long para poder ejecutarse en 32 bits
Dim hMenu As Long, hSubMenu As Long, R As Integer
Dim menRect As Rect
    If Button And 2 Then
        ScaleMode = TWIPS
        menRect.Left = 0
        menRect.Top = 0
        menRect.Right = Screen.Width / Screen.TwipsPerPixelX
        menRect.Bottom = Screen.Height / Screen.TwipsPerPixelY
        IX = (X + Left + ImgIcon(Index).Left) \ Screen.TwipsPerPixelX
        IY = (Y + Top + ImgIcon(Index).Top + ImgIcon(Index).Height) \ Screen.TwipsPerPixelY
        hMenu = GetMenu(frmMenus.hWnd)
        hSubMenu = GetSubMenu(hMenu, Index)
        R = TrackPopupMenu(hSubMenu, 2, IX, IY, 0, frmMenus.hWnd, menRect)
    End If

    ' Actualizar la información del sistema
    If Index = 2 Then
        FillSysInfo
    End If
End Sub

Private Sub tmrBounce_Timer()
'Lo siguiente es sólo texto estático para mejorar la velocidad
Static NewX As Integer, NewY As Integer, temp As Integer
Static hDC As Long, releaseit As Integer

'Calcular la nueva posición
    ScaleMode = PIXELS
    temp = X + dx
    If temp + PicWidth \ 2 > RightEdge Then
        dx = -Abs(dx)
    ElseIf temp < 0 Then
        dx = Abs(dx)
    End If
    
    NewX = X + dx
    
    temp = Y + dy
    If temp + PicHeight \ 2 > BottomEdge Then
        dy = -Abs(dy)
    ElseIf temp < 0 Then
        dy = Abs(dy)
    End If
    
    NewY = Y + dy

    If WindowState = MINIMIZED Then
        hDC = GetDC(GetDesktopWindow())
        releaseit = True
    Else
        hDC = Me.hDC
        releaseit = False
    End If


'Ahora realiza una llamada "transparente" a BitBlt:
'1 Copia la parte posterior del antiguo segundo plano sobre la antigua
'posición del gráfico
'2 Copia el segundo plano a donde va a ir el gráfico
'3 Dibuja la máscara
    temp = BitBlt(hDC, X, Y, PicWidth, PicHeight, picCopy.hDC, 0, 0, SRCCOPY)
    temp = BitBlt(picCopy.hDC, 0, 0, PicWidth, PicHeight, hDC, NewX, NewY, SRCCOPY)
    temp = BitBlt(hDC, NewX, NewY, PicWidth, PicHeight, picMask.hDC, 0, 0, SRCAND)
    temp = BitBlt(hDC, NewX, NewY, PicWidth, PicHeight, picSprite.hDC, 0, 0, SRCINVERT)
    X = NewX
    Y = NewY
    If releaseit Then ReleaseDC GetDesktopWindow(), hDC
End Sub

Private Sub tmrSysInfo_Timer()
Dim YourMemory As MEMORYSTATUS
Dim intX As Integer
Dim lWidth As Integer

    If fraInfo(RES_INFO).Visible Then
        For intX = 1 To 3
            lblR(intX).Visible = True
            lblResInfo(intX).Visible = True
            shpBar(intX).Visible = True
            shpFrame(intX).Visible = True
        Next intX
    Else
        For intX = 1 To 3
            lblR(intX).Visible = False
            lblResInfo(intX).Visible = False
            shpBar(intX).Visible = False
            shpFrame(intX).Visible = False
        Next intX
    End If
    
    YourMemory.dwLength = Len(YourMemory)
    GlobalMemoryStatus YourMemory
        
    With YourMemory
        lblInfo(6).Caption = "Memoria libre disponible (KB) : " & (.dwAvailPhys / 1024)
        
'       Comprueba la anchura antes de establecerla para probar y reducir
'       la intermitencia de la pantalla
        lWidth = shpFrame(1).Width * (.dwAvailPhys / .dwTotalPhys)
        If lWidth <> shpBar(1).Width Then
            shpBar(1).Width = lWidth
        End If
        
        lWidth = shpFrame(2).Width * (.dwAvailVirtual / .dwTotalVirtual)
        If lWidth <> shpBar(2).Width Then
            shpBar(2).Width = lWidth
        End If
        
        lWidth = shpFrame(3).Width * (.dwAvailPageFile / .dwTotalPageFile)
        If lWidth <> shpBar(3).Width Then
            shpBar(3).Width = lWidth
        End If
    End With
End Sub

