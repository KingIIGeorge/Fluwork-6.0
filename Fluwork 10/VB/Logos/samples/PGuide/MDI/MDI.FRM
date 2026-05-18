VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Bloc de notas MDI"
   ClientHeight    =   3495
   ClientLeft      =   915
   ClientTop       =   2205
   ClientWidth     =   5520
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   5460
      TabIndex        =   0
      Top             =   0
      Width           =   5520
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1995
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327680
         CancelError     =   -1  'True
         DefaultExt      =   "TXT"
         Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
         FilterIndex     =   557
         FontSize        =   1.1749e-38
      End
      Begin VB.Image imgPasteButtonUp 
         Height          =   330
         Left            =   5280
         Picture         =   "mdi.frx":0000
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgPasteButtonDn 
         Height          =   330
         Left            =   4920
         Picture         =   "mdi.frx":01E2
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCopyButtonDn 
         Height          =   330
         Left            =   4200
         Picture         =   "mdi.frx":03C4
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCopyButtonUp 
         Height          =   330
         Left            =   4560
         Picture         =   "mdi.frx":05A6
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCutButtonDn 
         Height          =   330
         Left            =   3840
         Picture         =   "mdi.frx":0788
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCutButtonUp 
         Height          =   330
         Left            =   3480
         Picture         =   "mdi.frx":096A
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgFileOpenButtonDn 
         Height          =   330
         Left            =   2760
         Picture         =   "mdi.frx":0B4C
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileOpenButtonUp 
         Height          =   330
         Left            =   3120
         Picture         =   "mdi.frx":0CD6
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileNewButtonUp 
         Height          =   330
         Left            =   2400
         Picture         =   "mdi.frx":0E60
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileNewButtonDn 
         Height          =   330
         Left            =   2040
         Picture         =   "mdi.frx":0FEA
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgPasteButton 
         Height          =   330
         Left            =   1560
         Picture         =   "mdi.frx":11CC
         ToolTipText     =   "Pegar"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgCopyButton 
         Height          =   330
         Left            =   1200
         Picture         =   "mdi.frx":13AE
         ToolTipText     =   "Copiar"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgCutButton 
         Height          =   330
         Left            =   840
         Picture         =   "mdi.frx":1590
         ToolTipText     =   "Cortar"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgFileOpenButton 
         Height          =   330
         Left            =   360
         Picture         =   "mdi.frx":1772
         ToolTipText     =   "Abrir archivo"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFileNewButton 
         Height          =   330
         Left            =   0
         Picture         =   "mdi.frx":18FC
         ToolTipText     =   "Archivo nuevo"
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOptionsToolbar 
         Caption         =   "&Barra de herramientas"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Formulario MDI principal de la aplicación de ejemplo       ***
'*** Bloc de notas MDI.                                         ***
'******************************************************************
Option Explicit
Private Sub imgCopyButton_Click()
    ' Actualiza la imagen.
    imgCopyButton.Refresh
    ' Llama al procedimiento de copiar
    EditCopyProc
End Sub

Private Sub imgCopyButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgCopyButton.Picture = imgCopyButtonDn.Picture
End Sub

Private Sub imgCopyButton_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgCopyButton.Width Or Y < 0 Or Y > imgCopyButton.Height Then
            imgCopyButton.Picture = imgCopyButtonUp.Picture
        Else
            imgCopyButton.Picture = imgCopyButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgCopyButton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgCopyButton.Picture = imgCopyButtonUp.Picture
End Sub

Private Sub imgCutButton_Click()
    ' Actualiza la imagen.
    imgCutButton.Refresh
    ' Llama al procedimiento de cortar
    EditCutProc
End Sub

Private Sub imgCutButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgCutButton.Picture = imgCutButtonDn.Picture
End Sub

Private Sub imgCutButton_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
Select Case Button
    Case 1
        If x <= 0 Or x > imgCutButton.Width Or Y < 0 Or Y > imgCutButton.Height Then
            imgCutButton.Picture = imgCutButtonUp.Picture
        Else
            imgCutButton.Picture = imgCutButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgCutButton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgCutButton.Picture = imgCutButtonUp.Picture
End Sub

Private Sub imgFileNewButton_Click()
    ' Actualiza la imagen.
    imgFileNewButton.Refresh
    ' Llama al procedimiento de nuevo archivo
    FileNew
End Sub

Private Sub imgFileNewButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgFileNewButton.Picture = imgFileNewButtonDn.Picture
End Sub

Private Sub imgFileNewButton_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgFileNewButton.Width Or Y < 0 Or Y > imgFileNewButton.Height Then
            imgFileNewButton.Picture = imgFileNewButtonUp.Picture
        Else
            imgFileNewButton.Picture = imgFileNewButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgFileNewButton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgFileNewButton.Picture = imgFileNewButtonUp.Picture
End Sub

Private Sub imgFileOpenButton_Click()
    ' Actualiza la imagen.
    imgFileOpenButton.Refresh
    ' Llama al procedimiento de abrir archivo
    FileOpenProc
End Sub

Private Sub imgFileOpenButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
End Sub

Private Sub imgFileOpenButton_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgFileOpenButton.Width Or Y < 0 Or Y > imgFileOpenButton.Height Then
            imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture
        Else
            imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgFileOpenButton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture
End Sub

Private Sub imgPasteButton_Click()
    ' Actualiza la imagen.
    imgPasteButton.Refresh
    ' Llama al procedimiento de pegar
    EditPasteProc
End Sub

Private Sub imgPasteButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgPasteButton.Picture = imgPasteButtonDn.Picture
End Sub

Private Sub imgPasteButton_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgPasteButton.Width Or Y < 0 Or Y > imgPasteButton.Height Then
            imgPasteButton.Picture = imgPasteButtonUp.Picture
        Else
            imgPasteButton.Picture = imgPasteButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgPasteButton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgPasteButton.Picture = imgPasteButtonUp.Picture
End Sub

Private Sub MDIForm_Load()
    ' La aplicación empieza aquí (evento Load del formulario Startup).
    Show
    ' Siempre establece el directorio de trabajo al directorio que contiene la aplicación.
    ChDir App.Path
    ' Inicializa la matriz de formularios Document y muestra el primer documento.
    ReDim Document(1)
    ReDim FState(1)
    Document(1).Tag = 1
    FState(1).Dirty = False
    ' Lee el Registro del sistema y establece la matriz de controles menú de archivos recientes del menú Archivo.
    GetRecentFiles
    ' Establece la variable pública gFindDirection, que determina la dirección
    ' de búsqueda de la función FindIt.
    gFindDirection = 1
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ' Si no se canceló el evento Unload (en los eventos QueryUnload de los formularios del Bloc de notas),
    ' no quedarán ventanas de documento, de modo que termina la aplicación.
    If Not AnyPadsLeft() Then
        End
    End If
End Sub

Private Sub mnuFileExit_Click()
    ' Termina la aplicación.
    End
End Sub

Private Sub mnuFileNew_Click()
    ' Llama al procedimiento de nuevo archivo
    FileNew
End Sub

Private Sub mnuFileOpen_Click()
    ' Llama al procedimiento de abrir archivo.
    FileOpenProc
End Sub

Private Sub mnuOptions_Click()
    ' Alterna la visibilidad de las barras de herramientas.
    mnuOptionsToolbar.Checked = frmMDI.picToolbar.Visible
End Sub

Private Sub mnuOptionsToolbar_Click()
    ' Llama al procedimiento de la barra de herramientas, pasando una referencia
    ' a este formulario.
    OptionsToolbarProc Me
End Sub

Private Sub mnuRecentFile_Click(index As Integer)
    ' Llama al procedimiento de abrir archivo, pasando una referencia
    ' al nombre de archivo seleccionado
    OpenFile (mnuRecentFile(index).Caption)
    ' Actualiza la lista de los archivos más recientemente abiertos.
    GetRecentFiles
End Sub

