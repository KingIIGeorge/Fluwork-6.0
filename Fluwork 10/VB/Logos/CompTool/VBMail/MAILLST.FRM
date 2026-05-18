VERSION 5.00
Begin VB.Form MailLst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lista de Mail"
   ClientHeight    =   3555
   ClientLeft      =   1320
   ClientTop       =   1785
   ClientWidth     =   6630
   Icon            =   "MAILLST.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   6630
   Begin VB.PictureBox Tools 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6630
      TabIndex        =   2
      Top             =   2940
      Width           =   6624
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   6660
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   6660
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image PrtImage 
         Height          =   480
         Left            =   1200
         Picture         =   "MAILLST.frx":030A
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Trash 
         Height          =   480
         Left            =   300
         Picture         =   "MAILLST.frx":0614
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.ListBox MList 
      Height          =   2040
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Headings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cabecera del cuadro de lista"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1680
   End
   Begin VB.Image NewMail 
      Height          =   480
      Left            =   5880
      Picture         =   "MAILLST.frx":091E
      Top             =   2820
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image nonew 
      Height          =   480
      Left            =   5280
      Picture         =   "MAILLST.frx":0C28
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "MailLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable a nivel de módulo para mantener la información de MouseDown.
Dim ListX, ListY

Private Sub Form_Load()
    ' Cambia el tamaño del formulario.
    Height = 3945
    Call Tools_Resize

     ' Establecer cabeceras del cuadro de lista.
     a$ = Mid$(Format$("De", "!" + String$(25, "@")), 1, 25)
     b$ = Mid$(Format$("Asunto", "!" + String$(35, "@")), 1, 35)
     c$ = "Recibido el"
     Headings = a$ + b$ + c$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Si el usuario está cerrando la aplicación, libera este formulario.
    If UnloadMode = 4 Then
        ' Está permitido liberar.
    Else
        ' Si el usuario sigue conectado, minimiza el formulario en lugar de cerrarlo.
        If VBMail.MapiMess.SessionID <> 0 Then
            Me.WindowState = 1
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    ' Si el formulario no está minimizado, cambia el tamaño del
    ' cuadro de lista para ajustarse al formulario.
    If WindowState <> 1 Then

        If VBMail.DispTools.Checked Then
            xHeight% = Tools.Height
        Else
            xHeight% = 0
        End If

        ' Buscar la altura mínima del formulario.
        If Height < 2500 - xHeight% Then
            Height = 2500
            Exit Sub
        End If

        MList.Width = ScaleWidth - MList.Left - 90
        MList.Height = ScaleHeight - 90 - MList.Top - xHeight%
    End If
End Sub

Private Sub MList_Click()
' Establece el índice de mensajes y habilita
' los botones "Anterior" y "Siguiente" según se haga necesario.
    Select Case MList.ListIndex
        Case 0
            VBMail.Previous.Enabled = False
        Case MList.ListCount - 1
            VBMail.Next.Enabled = False
        Case Else
            VBMail.Previous.Enabled = True
            VBMail.Next.Enabled = True
    End Select
    VBMail.MapiMess.MsgIndex = MList.ListIndex
End Sub

Private Sub MList_DBLClick()
' Verifica si el mensaje es visible actualmente,
' y si no lo es, lo carga dentro de un nuevo formulario.
    If Not MailLst.MList.ItemData(MailLst.MList.ListIndex) Then
       Dim Msg As New MsgView
       Call LoadMessage(MailLst.MList.ListIndex, Msg)
       MailLst.MList.ItemData(MailLst.MList.ListIndex) = True
    Else
        ' Busca en las ventanas activas para
        ' encontrar la ventana con el mensaje correcto a visualizar.
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) = MailLst.MList.ListIndex Then
                    Forms(i).Show
                    Exit Sub
                End If
            End If
        Next i
     End If
End Sub

Private Sub MList_KeyPress(KeyAscii As Integer)
    ' Si el usuario presiona ENTRAR, procesar la acción como un
    ' evento de doble clic.
    If KeyAscii = 13 Then
        Call MList_DBLClick
    End If
End Sub

Private Sub MList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Guardar los valores de X e Y para determinar
    ' el comienzo de la acción arrastrar y colocar.
    ListX = X
    ListY = Y
End Sub

Private Sub MList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón del ratón está presionado y la posición X,Y ha cambiado, comienza
    ' la acción de arrastrar.
    If Button = 1 And ((X <> ListX) Or (Y <> ListY)) Then
        MList.Drag 1
    End If
End Sub

Private Sub PrtImage_DragDrop(Source As Control, X As Single, Y As Single)
    ' Igual que File.PrintMessage en el menú "Archivo" de VBMAIL.
    Call PrintMail
End Sub

Private Sub Tools_Resize()
    ' Ajusta el ancho de las líneas en el frente de la barra de herramientas.
    Line1(0).X2 = Tools.Width
    Line1(1).X2 = Tools.Width
    Tools.Refresh
End Sub

Private Sub Trash_DragDrop(Source As Control, X As Single, Y As Single)
    ' Elimina un mensaje (Botón "Eliminar" o Edit.Delete).
   Call DeleteMessage
End Sub

