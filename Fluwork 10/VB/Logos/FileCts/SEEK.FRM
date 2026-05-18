VERSION 5.00
Begin VB.Form WinSeek 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinSeek"
   ClientHeight    =   4020
   ClientLeft      =   1920
   ClientTop       =   1890
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3720
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   720
      ScaleHeight     =   2895
      ScaleWidth      =   3855
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox lstFoundFiles 
         Height          =   2235
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lblfound 
         Caption         =   "&Archivos encontrados:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   3735
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.DriveListBox drvList 
         Height          =   1530
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.DirListBox dirList 
         Height          =   1695
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.FileListBox filList 
         Height          =   2040
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtSearchSpec 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "*.*"
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblCriteria 
         Caption         =   "&Criterio de búsqueda:"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   720
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Salir"
      Height          =   720
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "WinSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchFlag As Integer   ' Se usa como indicador para cancelar y otras operaciones.

Private Sub cmdExit_Click()
    If cmdExit.Caption = "&Salir" Then
        End
    Else                    ' Si el usuario eligió Cancelar, termina la búsqueda.
        SearchFlag = False
    End If
End Sub

Private Sub cmdSearch_Click()
' Inicializa la búsqueda y después realiza una búsqueda recursiva.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Comprueba lo que hizo el usuario en último lugar.
    If cmdSearch.Caption = "&Volver" Then  ' Si es restablecer, inicializa y sale.
        ResetSearch
        txtSearchSpec.SetFocus
        Exit Sub
    End If

    ' Actualiza dirList.Path si es distinto del directorio seleccionado
    ' actualmente; de lo contrario, realiza la búsqueda.
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub         ' Sale de forma que el usuario pueda mirar antes de buscar.
    End If

    ' Continúa con la búsqueda.
    Picture2.Move 0, 0
    Picture1.Visible = False
    Picture2.Visible = True

    cmdExit.Caption = "Cancelar"

    filList.Pattern = txtSearchSpec.Text
    FirstPath = dirList.Path
    DirCount = dirList.ListCount

    ' Inicia la búsqueda recursiva de directorios.
    NumFiles = 0     ' Restablece el indicador de archivos encontrados.
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
    cmdSearch.Caption = "&Volver"
    cmdSearch.SetFocus
    cmdExit.Caption = "&Salir"
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Busca recursivamente en directorios desde NewPath hacia abajo...
'  Se busca en NewPath en este paso recursivo.
'  BackUp es el origen de este paso recursivo.
'  DirCount es el número de subdirectorios de este directorio.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True           ' Establece el indicador de modo que el usuario pueda interrumpir.
    DirDiver = False            ' Lo establece a True si hay un error.
    retval = DoEvents()         ' Comprueba si hay eventos (por ejemplo, si el usuario elige Cancelar).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount                  ' ¿Cuántos directorios hay debajo de éste?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path                      ' Guarda la ruta de acceso anterior para el próximo paso recursivo.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Obtiene hasta la parte inferior del nodo.
            dirList.Path = dirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Sube un nivel en los directorios.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Llama a una función para enumerar archivos.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then              ' Comprueba 2 bytes/carácter
            ThePath = dirList.Path                  ' Si está a nivel raíz, lo deja como está...
        Else
            ThePath = dirList.Path + "\"            ' De lo contrario, pone "\" delante del nombre de archivo.
        End If
        For ind = 0 To filList.ListCount - 1        ' Agrega archivos de este directorio al cuadro de lista.
            entry = ThePath + filList.List(ind)
            lstFoundFiles.AddItem entry
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then        ' Si hay un directorio superior, va a él.
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' Si ocurre un error por falta de memoria, supone que el cuadro de lista se llenó.
        DirDiver = True         ' Crea el mensaje y establece el valor de retorno AbandonSearch.
        MsgBox "Se ha llenado el cuadro de lista. Abandonando la búsqueda..."
        Exit Function           ' Observe que el procedimiento de salida restablece Err a 0.
    Else                        ' De lo contrario, muestra un mensaje de error y sale.
        MsgBox Error
        End
    End If
End Function

Private Sub DirList_Change()
    ' Actualiza el cuadro de lista de archivos para sincronizar con el cuadro de lista de directorios.
    filList.Path = dirList.Path
End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub Form_Load()
    Picture2.Move 0, 0
    Picture2.Width = WinSeek.ScaleWidth
    Picture2.BackColor = WinSeek.BackColor
    lblCount.BackColor = WinSeek.BackColor
    lblCriteria.BackColor = WinSeek.BackColor
    lblfound.BackColor = WinSeek.BackColor
    Picture1.Move 0, 0
    Picture1.Width = WinSeek.ScaleWidth
    Picture1.BackColor = WinSeek.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub ResetSearch()
    ' Reinicializa antes de iniciar una nueva búsqueda.
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False        ' Indicador de que hay una búsqueda en curso.
    Picture2.Visible = False
    cmdSearch.Caption = "&Buscar"
    cmdExit.Caption = "&Salir"
    Picture1.Visible = True
    dirList.Path = CurDir: drvList.Drive = dirList.Path ' Restablece la ruta de acceso.
End Sub

Private Sub txtSearchSpec_Change()
    ' Actualiza el cuadro de lista de archivos si el usuario cambia el modelo.
    filList.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0          ' Resalta la entrada actual.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub

