VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmSDI 
   Caption         =   "Bloc de notas SDI - Sintítulo"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   810
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6690
   Begin VB.TextBox txtNote 
      Height          =   3975
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
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
      ScaleWidth      =   6630
      TabIndex        =   1
      Top             =   0
      Width           =   6690
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1995
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         CancelError     =   -1  'True
         DefaultExt      =   "TXT"
         Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
         FilterIndex     =   473
         FontSize        =   1.1749e-38
      End
      Begin VB.Image imgFileNewButton 
         Height          =   330
         Left            =   0
         Picture         =   "frmsdi.frx":0000
         ToolTipText     =   "Nuevo archivo"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFileOpenButton 
         Height          =   330
         Left            =   360
         Picture         =   "frmsdi.frx":018A
         ToolTipText     =   "Abrir archivo"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgCutButton 
         Height          =   330
         Left            =   840
         Picture         =   "frmsdi.frx":0314
         ToolTipText     =   "Cortar"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgCopyButton 
         Height          =   330
         Left            =   1200
         Picture         =   "frmsdi.frx":04F6
         ToolTipText     =   "Copiar"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgPasteButton 
         Height          =   330
         Left            =   1560
         Picture         =   "frmsdi.frx":06D8
         ToolTipText     =   "Pegar"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgFileNewButtonDn 
         Height          =   330
         Left            =   2040
         Picture         =   "frmsdi.frx":08BA
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgFileNewButtonUp 
         Height          =   330
         Left            =   2400
         Picture         =   "frmsdi.frx":0A9C
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileOpenButtonUp 
         Height          =   330
         Left            =   3120
         Picture         =   "frmsdi.frx":0C26
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileOpenButtonDn 
         Height          =   330
         Left            =   2760
         Picture         =   "frmsdi.frx":0DB0
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgCutButtonUp 
         Height          =   330
         Left            =   3480
         Picture         =   "frmsdi.frx":0F3A
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCutButtonDn 
         Height          =   330
         Left            =   3840
         Picture         =   "frmsdi.frx":111C
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCopyButtonUp 
         Height          =   330
         Left            =   4560
         Picture         =   "frmsdi.frx":12FE
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCopyButtonDn 
         Height          =   330
         Left            =   4200
         Picture         =   "frmsdi.frx":14E0
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgPasteButtonDn 
         Height          =   330
         Left            =   4920
         Picture         =   "frmsdi.frx":16C2
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgPasteButtonUp 
         Height          =   330
         Left            =   5280
         Picture         =   "frmsdi.frx":18A4
         Top             =   0
         Visible         =   0   'False
         Width           =   375
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
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Guardar &como"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
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
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cor&tar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Eliminar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Seleccionar todo"
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "&Hora y fecha"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Buscar"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Buscar &siguiente"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOptionsToolbar 
         Caption         =   "&Barra de herramientas"
      End
      Begin VB.Menu mnuFonts 
         Caption         =   "&Fuentes"
         Begin VB.Menu mnuFontName 
            Caption         =   "Nombre de la fuente"
            Index           =   0
         End
      End
      Begin VB.Menu mnuOptionsLaunch 
         Caption         =   "&Iniciar nueva instancia"
      End
   End
End
Attribute VB_Name = "frmSDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Formulario principal de la aplicación de ejemplo Bloc de notas SDI  ***
'***************************************************************************
Option Explicit

Private Sub Form_Load()
    Dim i As Integer        ' Variable contador.
    
    ' La aplicación empieza aquí (evento Load del formulario Startup).
    Show
    ' Siempre establece el directorio de trabajo al directorio que contiene la aplicación.
    ChDir App.Path
    FState.Dirty = False
    ' Lee el Registro del sistema y establece la lista de archivos recientes del menú Archivo.
    GetRecentFiles
    ' Establece la variable pública gFindDirection, que determina la dirección de búsqueda
    ' de la función FindIt.
    gFindDirection = 1
        
    ' Asigna el nombre de la primera fuente a una entrada
    ' del menú de fuentes, después recorre la colección de las fuentes,
    ' agregándolas al menú
    mnuFontName(0).Caption = Screen.Fonts(0)
    For i = 1 To Screen.FontCount - 1
        Load mnuFontName(i)
        mnuFontName(0).Caption = Screen.Fonts(i)
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    Dim strFilename As String
    Dim intResponse As Integer

    ' Comprueba si el texto ha sido modificado.
    If FState.Dirty Then
        strFilename = Me.Caption
        strMsg = "El texto en [" & strFilename & "] ha cambiado."
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "¿Desea guardar los cambios?"
        intResponse = MsgBox(strMsg, 51, frmSDI.Caption)
        Select Case intResponse
            Case 6      ' El usuario eligió Sí.
                If Left(Me.Caption, 8) = "Sintítulo" Then
                    ' El archivo no ha sido guardado aún.
                    strFilename = "sintítulo.txt"
                    ' Obtiene strFilename, y después llama al procedimiento de guardar.
                    strFilename = GetFileName(strFilename)
                Else
                    ' El título del formulario contiene el nombre del archivo abierto.
                    strFilename = Me.Caption
                End If
                ' Llama al procedimiento de guardar. Si strFilename está vacía, es porque
                ' el usuario eligió Cancelar en el cuadro de diálogo Guardar como;
                ' si no, guarda el archivo.
                If strFilename <> "" Then
                    SaveFileAs strFilename
                End If
            Case 7      ' El usuario eligió No. Descarga el archivo.
                Cancel = False
            Case 2      ' El usuario eligió Cancelar. Cancela la descarga.
                Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Resize()
    ' Llama al procedimiento de cambiar el tamaño
    ResizeNote
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Llama al procedimiento de enumerar los archivos más recientes
    GetRecentFiles
End Sub

Private Sub imgCopyButton_Click()
    ' Actualiza la imagen.
    imgCopyButton.Refresh
    ' Llama al procedimiento de copiar
    EditCopyProc
End Sub

Private Sub imgCopyButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgCopyButton.Picture = imgCopyButtonDn.Picture
End Sub

Private Sub imgCopyButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgCopyButton.Width Or Y < 0 Or Y > imgCopyButton.Height Then
            imgCopyButton.Picture = imgCopyButtonUp.Picture
        Else
            imgCopyButton.Picture = imgCopyButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgCopyButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgCopyButton.Picture = imgCopyButtonUp.Picture
End Sub

Private Sub imgCutButton_Click()
    ' Actualiza la imagen.
    imgCutButton.Refresh
    ' Llama al procedimiento de cortar
    EditCutProc
End Sub

Private Sub imgCutButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgCutButton.Picture = imgCutButtonDn.Picture
End Sub

Private Sub imgCutButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgCutButton.Width Or Y < 0 Or Y > imgCutButton.Height Then
            imgCutButton.Picture = imgCutButtonUp.Picture
        Else
            imgCutButton.Picture = imgCutButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgCutButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgCutButton.Picture = imgCutButtonUp.Picture
End Sub

Private Sub imgFileNewButton_Click()
    ' Actualiza la imagen.
    imgFileNewButton.Refresh
    ' Llama al procedimiento de nuevo archivo
    FileNew
End Sub

Private Sub imgFileNewButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgFileNewButton.Picture = imgFileNewButtonDn.Picture
End Sub

Private Sub imgFileNewButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgFileNewButton.Width Or Y < 0 Or Y > imgFileNewButton.Height Then
            imgFileNewButton.Picture = imgFileNewButtonUp.Picture
        Else
            imgFileNewButton.Picture = imgFileNewButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgFileNewButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgFileNewButton.Picture = imgFileNewButtonUp.Picture
End Sub

Private Sub imgFileOpenButton_Click()
    ' Actualiza la imagen.
    imgFileOpenButton.Refresh
    ' Llama al procedimiento de abrir archivo
    FileOpenProc
End Sub

Private Sub imgFileOpenButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
End Sub

Private Sub imgFileOpenButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgFileOpenButton.Width Or Y < 0 Or Y > imgFileOpenButton.Height Then
            imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture
        Else
            imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgFileOpenButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture

End Sub

Private Sub imgPasteButton_Click()
    ' Actualiza la imagen.
    imgPasteButton.Refresh
    ' Llama al procedimiento de pegar
    EditPasteProc
End Sub

Private Sub imgPasteButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgPasteButton.Picture = imgPasteButtonDn.Picture
End Sub

Private Sub imgPasteButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgPasteButton.Width Or Y < 0 Or Y > imgPasteButton.Height Then
            imgPasteButton.Picture = imgPasteButtonUp.Picture
        Else
            imgPasteButton.Picture = imgPasteButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgPasteButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgPasteButton.Picture = imgPasteButtonUp.Picture
End Sub

Private Sub mnuEditCopy_Click()
    ' Llama al procedimiento de copiar
    EditCopyProc
End Sub

Private Sub mnuEditCut_Click()
    ' Llama al procedimiento de cortar
    EditCutProc
End Sub

Private Sub mnuEditDelete_Click()
' Si el puntero del mouse no se encuentra al final del bloc de notas ...
    If txtNote.SelStart <> Len(Screen.ActiveControl.Text) Then
        ' Si no hay texto seleccionado, extiende la selección en uno.
        If txtNote.SelLength = 0 Then
            txtNote.SelLength = 1
            ' Si el puntero del mouse se encuentra en una línea en blanco, extiende la selección en dos.
            If Asc(txtNote.SelText) = 13 Then
                txtNote.SelLength = 2
            End If
        End If
        ' Elimina el texto seleccionado.
        txtNote.SelText = ""
    End If
End Sub

Private Sub mnuEditPaste_Click()
    ' Llama al procedimiento de pegar.
    EditPasteProc
End Sub

Private Sub mnuEditSelectAll_Click()
    ' Utiliza SelStart y SelLength para seleccionar el texto.
    txtNote.SelStart = 0
    txtNote.SelLength = Len(txtNote.Text)
End Sub

Private Sub mnuEditTime_Click()
    ' Inserta la hora y la fecha actual.
    txtNote.SelText = Now
End Sub

Private Sub mnuFileExit_Click()
    ' Termina la aplicación.
    Unload Me
End Sub

Public Sub mnuFileNew_Click()
    ' Llama al procedimiento de nuevo formulario
    FileNew
End Sub

Private Sub mnuFileOpen_Click()
    ' Llama al procedimiento de abrir archivo.
    FileOpenProc
End Sub

Private Sub mnuFileSave_Click()
    ' Llama al procedimiento de guardar archivo
    FileSave
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    ' Asigna el título del formulario a la variable.
    strDefaultName = Right$(Me.Caption, Len(Me.Caption) - 20)
    If Me.Caption = "Bloc de notas SDI - Sintítulo" Then
        ' El archivo no ha sido guardado aún.
        ' Obtiene el nombre del archivo y después llama al procedimiento de guardar.
        
        strSaveFileName = GetFileName("sintítulo.txt")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Actualiza la lista de archivos más recientes del menú Archivo.
        UpdateFileMenu (strSaveFileName)
    Else
        ' El título del formulario contiene el nombre del archivo abierto.
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Actualiza la lista de archivos más recientes del menú Archivo.
        UpdateFileMenu (strSaveFileName)
    End If
End Sub

Private Sub mnuFontName_Click(Index As Integer)
    ' Asigna la fuente seleccionada a la propiedad fuente del cuadro de texto.
    txtNote.FontName = mnuFontName(Index).Caption
End Sub

Private Sub mnuOptions_Click()
    ' Alterna la propiedad Checked para que coincida con la propiedad Visible.
    mnuOptionsToolbar.Checked = picToolbar.Visible
End Sub

Private Sub mnuOptionsLaunch_Click()
    Dim strApp As String
    
    ' Inicia una nueva instancia del bloc de notas.
    strApp = App.Path & "\" & App.EXEName
    Shell strApp, 1
End Sub

Private Sub mnuOptionsToolbar_Click()
    ' Alterna la propiedad visible de la barra de herramientas
    picToolbar.Visible = Not picToolbar.Visible
    ' Cambia la marca para que coincida con el estado actual
    mnuOptionsToolbar.Checked = picToolbar.Visible
    ' Llama al procedimiento de cambiar el tamaño
    ResizeNote
End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
    ' Llama al procedimiento de abrir archivos, pasando una
    ' referencia al nombre del archivo seleccionado
    OpenFile (mnuRecentFile(Index).Caption)
    ' Actualiza la lista de archivos más recientes del menú Archivo.
    GetRecentFiles
End Sub

Private Sub mnuSearchFind_Click()
    ' Si el cuadro de texto contiene algo de texto, lo asigna al
    ' cuadro de texto del formulario Buscar, si no, asigna
    ' el último valor de búsqueda.
    If txtNote.SelText <> "" Then
        frmFind.txtFind.Text = txtNote.SelText
    Else
        frmFind.txtFind.Text = gFindString
    End If
    ' Establece la variable pública de inicio al principio.
    gFirstTime = True
    ' Establece la casilla de verificación de distinción entre mayúsculas y minúsculas
    ' de acuerdo con la variable pública
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Presenta el formulario Buscar.
    frmFind.Show vbModal
End Sub

Private Sub mnuSearchFindNext_Click()
    ' Si la variable pública no está vacía, llama al
    ' procedimiento de búsqueda, si no, llama al menú de búsqueda
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSearchFind_Click
    End If
End Sub

Private Sub txtNote_Change()
    ' Establece la variable pública para mostrar que el texto ha sido modificado.
    FState.Dirty = True
End Sub
