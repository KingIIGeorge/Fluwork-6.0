VERSION 5.00
Begin VB.Form frmNotePad 
   Caption         =   "Sinnombr"
   ClientHeight    =   3990
   ClientLeft      =   1515
   ClientTop       =   3315
   ClientWidth     =   5670
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3990
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   101.07
   Begin VB.TextBox Text1 
      Height          =   3975
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "C&errar"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Guardar &como..."
      End
      Begin VB.Menu mnuFSep 
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
         Caption         =   "C&ortar"
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
         Caption         =   "Eli&minar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuESep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Seleccionar &todo"
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "&Hora/Fecha"
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
      Begin VB.Menu mnuFont 
         Caption         =   "&Fuente"
         Begin VB.Menu mnuFontName 
            Caption         =   "FontName"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Mosaico"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Organizar iconos"
      End
   End
End
Attribute VB_Name = "frmNotePad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Formulario hijo de la aplicación de ejemplo Bloc de notas MDI   ***
'***********************************************************************
Option Explicit

Private Sub Form_Load()
    Dim i As Integer        ' Variable contador.
    
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
    If FState(Me.Tag).Dirty Then
        strFilename = Me.Caption
        strMsg = "El texto en [" & strFilename & "] ha cambiado."
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "¿Desea guardar los cambios?"
        intResponse = MsgBox(strMsg, 51, frmMDI.Caption)
        Select Case intResponse
            Case 6      ' El usuario eligió Sí.
                If Left(Me.Caption, 8) = "Sinnombr" Then
                    ' El archivo no ha sido guardado aún.
                    strFilename = "sinnombr.txt"
                    ' Obtiene strFilename y después llama al procedimiento de guardar.
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
    ' Expande el cuadro de texto para que ocupe el área interna del formulario hijo actual.
    Text1.Height = ScaleHeight
    Text1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Muestra la instancia del formulario actual como eliminada
    FState(Me.Tag).Deleted = True
    
    ' Oculta los botones de edición de la barra de herramientas si no hay ventanas en el bloc de notas.
    If Not AnyPadsLeft() Then
        frmMDI.imgCutButton.Visible = False
        frmMDI.imgCopyButton.Visible = False
        frmMDI.imgPasteButton.Visible = False
        ' Alterna la variable pública de la barra de herramientas
        gToolsHidden = True
        ' Llama al procedimiento de la lista de archivos más recientes
        GetRecentFiles
    End If
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
    If Screen.ActiveControl.SelStart <> Len(Screen.ActiveControl.Text) Then
        ' Si no hay texto seleccionado, extiende la selección en uno.
        If Screen.ActiveControl.SelLength = 0 Then
            Screen.ActiveControl.SelLength = 1
            ' Si el puntero del mouse se encuentra en una línea en blanco, extiende la selección en dos.
            If Asc(Screen.ActiveControl.SelText) = 13 Then
                Screen.ActiveControl.SelLength = 2
            End If
        End If
        ' Elimina el texto seleccionado.
        Screen.ActiveControl.SelText = ""
    End If
End Sub

Private Sub mnuEditPaste_Click()
    ' Llama al procedimiento de pegar.
    EditPasteProc
End Sub

Private Sub mnuEditSelectAll_Click()
    ' Utiliza SelStart y SelLength para seleccionar el texto.
    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.Text)
End Sub

Private Sub mnuEditTime_Click()
    ' Inserta la hora y la fecha actual.
    Text1.SelText = Now
End Sub

Private Sub mnuFileClose_Click()
    ' Descarga este formulario.
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    ' Al descargar el formulario MDI se invoca el evento QueryUnload
    ' de cada formulario hijo, y después el del formulario MDI.
    ' Al establecer el argumento Cancel a True en cualquiera de los
    ' eventos QueryUnload se cancela la descarga.
    Unload frmMDI
End Sub

Private Sub mnuFileNew_Click()
    ' Llama al procedimiento de nuevo formulario
    FileNew
End Sub

Private Sub mnuFontName_Click(index As Integer)
    ' Asigna la fuente seleccionada a la propiedad fuente del cuadro de texto.
    Text1.FontName = mnuFontName(index).Caption
End Sub

Private Sub mnuFileOpen_Click()
    ' Llama al procedimiento de abrir archivo.
    FileOpenProc
End Sub

Private Sub mnuFileSave_Click()
    Dim strFilename As String

    If Left(Me.Caption, 8) = "Sinnombr" Then
        ' El archivo no ha sido guardado aún.
        ' Obtiene el nombre del archivo y después llama al procedimiento de guardar.
        strFilename = GetFileName(strFilename)
    Else
        ' El título del formulario contiene el nombre del archivo abierto.
        strFilename = Me.Caption
    End If
    ' Llama al procedimiento de guardar. Si Filename está vacía, es porque el
    ' usuario eligió Cancelar en el cuadro de diálogo Guardar como; si no,
    ' guarda el archivo.
    If strFilename <> "" Then
        SaveFileAs strFilename
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    ' Asigna el título del formulario a la variable.
    strDefaultName = Me.Caption
    If Left(Me.Caption, 8) = "Sinnombr" Then
        ' El archivo no ha sido guardado aún.
        ' Obtiene el nombre del archivo, y después llama al procedimiento de guardar.
        
        strSaveFileName = GetFileName("Sinnombr.txt")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Actualiza la lista de archivos más recientemente abiertos en el menú Archivo.
        UpdateFileMenu (strSaveFileName)
    Else
        ' El título del formulario contiene el nombre del archivo abierto.
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Actualiza la lista de archivos más recientemente abiertos en el menú Archivo.
        UpdateFileMenu (strSaveFileName)
    End If

End Sub

Private Sub mnuOptions_Click()
    ' Alterna la propiedad Checked para que coincida con la propiedad Visible.
    mnuOptionsToolbar.Checked = frmMDI.picToolbar.Visible
End Sub

Private Sub mnuOptionsToolbar_Click()
    ' Llama al procedimiento de la barra de herramientas, pasando una referencia
    ' a esta instancia de formulario.
    OptionsToolbarProc Me
End Sub

Private Sub mnuRecentFile_Click(index As Integer)
    ' Llama al procedimiento de abrir, pasando una
    ' referencia al nombre del archivo seleccionado
    OpenFile (mnuRecentFile(index).Caption)
    ' Actualiza la lista de archivos más recientemente abiertos en el menú Archivo.
    GetRecentFiles
End Sub

Private Sub mnuSearchFind_Click()
    ' Si hay texto en este cuadro de texto, lo asigna al
    ' cuadro de texto del formulario Find. Si no, asigna el valor de la
    ' última búsqueda.
    If Me.Text1.SelText <> "" Then
        frmFind.Text1.Text = Me.Text1.SelText
    Else
        frmFind.Text1.Text = gFindString
    End If
    ' Establece la variable pública de inicio al principio.
    gFirstTime = True
    ' Establece la casilla de verificación de distinción entre mayúsculas y minúsculas
    ' según la variable pública
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Presenta el formulario Buscar.
    frmFind.Show vbModal
End Sub

Private Sub mnuSearchFindNext_Click()
    ' Si la variable pública no está vacía, llama al
    ' procedimiento de búsqueda, si no llama al menú de búsqueda
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSearchFind_Click
    End If
End Sub

Private Sub mnuWindowArrange_Click()
    ' Organiza los iconos de los formularios hijos minimizados.
    frmMDI.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    ' Organiza los formularios hijos en cascada.
    frmMDI.Arrange vbCascade
End Sub

Private Sub mnuWindowTile_Click()
    ' Organiza los formularios hijos en mosaico.
    frmMDI.Arrange vbTileHorizontal
End Sub

Private Sub Text1_Change()
    ' Establece la variable pública para mostrar que el texto ha sido modificado.
    FState(Me.Tag).Dirty = True
End Sub

