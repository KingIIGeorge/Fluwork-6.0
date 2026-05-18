VERSION 5.00
Begin VB.Form frmDrag 
   Caption         =   "Arrastrar y colocar"
   ClientHeight    =   2670
   ClientLeft      =   2130
   ClientTop       =   3780
   ClientWidth     =   6405
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2670
   ScaleWidth      =   6405
   Begin VB.DriveListBox Drive1 
      DragIcon        =   "DRAG.frx":0000
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   2280
      Pattern         =   "*.txt;*.bmp;*.exe;*.hlp"
      TabIndex        =   1
      Top             =   120
      Width           =   2052
   End
   Begin VB.DirListBox Dir1 
      DragIcon        =   "DRAG.frx":030A
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveErrs
    Dir1.Path = Drive1.Drive
Exit Sub
DriveErrs:
    Select Case Err
        Case 68
            MsgBox prompt:="La unidad no está preparada. Inserte un disco en la unidad.", _
            buttons:=vbExclamation
            ' Restablece la ruta a la unidad anterior.
            Drive1.Drive = Dir1.Path
            Exit Sub
        Case Else
            MsgBox prompt:="Error en la aplicación.", buttons:=vbExclamation
    End Select
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    File1.DragIcon = Drive1.DragIcon
    File1.Drag
End Sub

Private Sub Form_Load()
    frmDrag.Width = 6525
    frmDrag.Height = 3075
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)

    ' Obtiene las tres últimas letras del nombre del archivo arrastrado.
    temp = Right$(File1.filename, 3)

    ' Si el archivo arrastrado se encuentra en la raíz, agrega el nombre del archivo.
    If Mid$(File1.Path, Len(File1.Path)) = "\" Then
      dropfile = File1.Path & File1.filename
    ' Si el archivo arrastrado no se encuentra en la raíz, agrega "\" al nombre del archivo.
    Else
      dropfile = File1.Path & "\" & File1.filename
    End If
      
    Image1.Picture = LoadPicture("")
    Select Case UCase$(Trim$(temp))
        Case "TXT"
            X = Shell("Notepad " + dropfile, 1)
        Case "BMP"
            Image1.Picture = LoadPicture(dropfile)
        Case "EXE"
            X = Shell(dropfile, 1)
        Case "HLP"
            X = Shell("WinHelp " + dropfile, 1)
        Case Else
            msg = "Pruebe con uno de estos tipos de archivos:"
            msg = vbCrLf & msg & vbCrLf & vbCrLf & "     .txt, .bmp, .exe, .hlp"
            MsgBox msg
    End Select

End Sub

Private Sub Image1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Select Case State
    Case 0
        ' Presenta un icono nuevo cuando el origen entra en el área de colocar.
        File1.DragIcon = Dir1.DragIcon
    Case 1
        ' Presenta el DragIcon original cuando el origen sale del área de colocar.
        File1.DragIcon = Drive1.DragIcon
    End Select

' Observe que Dir1.DragIcon y Drive1.DragIcon han sido establecidos
' en tiempo de diseño. Esto le permite cargar los iconos "Enter"
' y "Leave" de File1 en tiempo de ejecución sin que sea necesario
' que el usuario los tenga en el disco.
End Sub

