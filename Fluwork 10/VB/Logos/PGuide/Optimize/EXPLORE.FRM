VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmExplore 
   Caption         =   "Optimizar VB 5.0"
   ClientHeight    =   5865
   ClientLeft      =   1425
   ClientTop       =   900
   ClientWidth     =   3420
   ClipControls    =   0   'False
   Icon            =   "Explore.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   3420
   WhatsThisHelp   =   -1  'True
   Begin ComctlLib.TreeView tvExample 
      Height          =   5130
      Left            =   -30
      TabIndex        =   0
      Top             =   450
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   9049
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilExplore"
      BorderStyle     =   1
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5580
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   503
      SimpleText      =   ""
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2558
            MinWidth        =   2558
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "12:35 PM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/30/96"
            Object.Tag             =   ""
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
   Begin ComctlLib.ImageList ilExplore 
      Left            =   2715
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":0736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":0848
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explore.frx":0B7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   9195
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Label lblInfoWnd 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ejemplos de optimización de código"
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   105
      Width           =   3400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   9195
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&Archivo"
      Begin VB.Menu mnu_FileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&Ver"
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnu_ToolsCloseAll 
         Caption         =   "&Cerrar todas las ventanas"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnu_HelpAbout 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "frmExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmlastForm As Form
Private Sub Form_Load()
    
  Me.Move 0, (Screen.Height - Me.Height) / 2
  Me.Show
  DoEvents

End Sub


Private Sub Form_Paint()
  Debug.Print "frmExplore se ha dibujado"
End Sub

Private Sub Form_Resize()
  Dim lfrmScaleWidth As Long
  Dim lRetVal As Long
  
  On Error Resume Next
  ' Desactiva la actualización automática de la pantalla
  'lRetVal = SetRedraw(Me.hWnd, 0)
  
  lfrmScaleWidth = frmExplore.ScaleWidth
  
  tvExample.Move 0, 495, lfrmScaleWidth, frmExplore.ScaleHeight - tvExample.Top - StatusBar1.Height - 20
  'tvExample.Width = lfrmScaleWidth
  'tvExample.Height = frmExplore.ScaleHeight - tvExample.Top - StatusBar1.Height - 20
  
  lblInfoWnd(0).Width = lfrmScaleWidth
  Line1(0).X2 = lfrmScaleWidth
  Line1(1).X2 = lfrmScaleWidth
  On Error GoTo 0
  
  ' Activa la actualización automática de la pantalla 
  'lRetVal = SetRedraw(Me.hWnd, 1)

End Sub
Private Sub Form_Unload(Cancel As Integer)
  Dim iFormCount As Integer

  For iFormCount = Forms.Count - 1 To 0 Step -1
    Unload Forms(iFormCount)
  Next iFormCount

End Sub


Private Sub mnu_FileExit_Click()
  Unload Me
End Sub


Private Sub mnu_HelpAbout_Click()
  frmAbout.Show
End Sub


Private Sub mnu_ToolsCloseAll_Click()
  Dim iFormCount As Integer

  For iFormCount = Forms.Count - 1 To 1 Step -1
    Unload Forms(iFormCount)
  Next iFormCount
End Sub

Private Sub tvExample_NodeClick(ByVal Node As Node)
  Dim lRetVal As Long
  
  Select Case Node.Text
    Case "Algoritmos"
      Algorithms.Show
      Algorithms.SetFocus
    Case "Dibujar imagen"
      frmPaintPic.Show
      frmPaintPic.SetFocus
    Case "Gráficos"
      Graphics.Show
      Graphics.SetFocus
    Case "Cadenas y archivos"
      Strings.Show
      Strings.SetFocus
    Case "Optimización de código"
      Code.Show
      Code.SetFocus
    Case "Tipos numéricos"
      frmNumbers.Show
      frmNumbers.SetFocus
    
    Case "Picture frente a Image"
      Pictures.Show
      Pictures.SetFocus
    Case "Destruir objetos de formularios"
      frmShow2.Show
      frmShow2.SetFocus
    Case "Pantalla inicial"
      frmSplashDemo.Show
      frmSplashDemo.SetFocus
    Case "Colecciones"
      frmCollection.Show
      frmCollection.SetFocus
    Case Else
      MsgBox "No sé el tipo de formulario.", vbQuestion, "Mostrar un error de ejemplo"
    End Select
  DoEvents
End Sub


