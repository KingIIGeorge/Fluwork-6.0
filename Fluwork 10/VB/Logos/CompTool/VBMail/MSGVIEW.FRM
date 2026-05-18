VERSION 5.00
Begin VB.Form MsgView 
   Caption         =   "Mensaje Mail"
   ClientHeight    =   4380
   ClientLeft      =   1575
   ClientTop       =   1620
   ClientWidth     =   6210
   Icon            =   "MSGVIEW.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   6210
   Begin VB.PictureBox AttachWin 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6150
      TabIndex        =   12
      Top             =   3555
      Visible         =   0   'False
      Width           =   6210
      Begin VB.ListBox aList 
         Height          =   90
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   5955
      End
      Begin VB.Label NumAtt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   60
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datos adjuntos:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   60
         Width           =   1110
      End
   End
   Begin VB.TextBox txtNoteText 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1740
      Width           =   6195
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   6150
      TabIndex        =   6
      Top             =   0
      Width           =   6210
      Begin VB.TextBox txtsubject 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   5
         Top             =   1320
         Width           =   4995
      End
      Begin VB.TextBox txtcc 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   4
         Top             =   1020
         Width           =   4995
      End
      Begin VB.TextBox txtTo 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   3
         Top             =   720
         Width           =   4995
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   2
         Top             =   420
         Width           =   4995
      End
      Begin VB.TextBox txtFrom 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   1
         Top             =   120
         Width           =   4995
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H00FFFFFF&
         X1              =   6000
         X2              =   6000
         Y1              =   60
         Y2              =   1620
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H00FFFFFF&
         X1              =   60
         X2              =   6000
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00808080&
         X1              =   60
         X2              =   5940
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H00808080&
         X1              =   60
         X2              =   60
         Y1              =   60
         Y2              =   1620
      End
   End
End
Attribute VB_Name = "MsgView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aList_DblClick()
    ' ListIndex es el índice dentro de la lista de datos adjuntos.
    VBMail.MapiMess.AttachmentIndex = aList.ListIndex
    If VBMail.MapiMess.AttachmentType = vbAttachTypeData Then
        Call DisplayAttachedFile(VBMail.MapiMess.AttachmentPathName)
    Else
        MsgBox " Esta aplicación no visualiza objetos tipo OLE adjuntos"
    End If
End Sub

Private Sub AttachWin_Resize()
    '  Actualizar los anchos de los campos y ajustar los controles
    '  de línea cuando sea necesario.
    aList.Width = AttachWin.Width - aList.Left - 315
End Sub

Private Sub Form_Activate()
    ' Cuando se activa el formulario, actualizar mailLst.MList
    ' para reflejar el elemento actual. La propiedad Tag contiene
    ' el índice del mensaje que se muestra actualmente.
    MailLst.MList.ListIndex = Val(Me.Tag)
    MailLst.MList.ItemData(Val(Me.Tag)) = True
    VBMail.MapiMess.MsgIndex = Val(Me.Tag)
End Sub

Private Sub Form_Load()
    ' Asegura que todos los cambios de tamaño se hacen al comenzar.
    Call Picture1_Resize
    Call AttachWin_Resize
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    ' Ajusta el tamaño de la ventana si el formulario no está minimizado.
    Call SizeMessageWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Tag se establece a -1 después de que se haya eliminado el
    ' mensaje mostrado actualmente.
    If Val(Me.Tag) > 0 Then
        MailLst.MList.ItemData(Val(Me.Tag)) = False
    End If
End Sub

Private Sub Picture1_Resize()
    '  Actualiza los anchos de los campos y ajusta los controles de línea
    '  cuando es necesario.
    TopLine.X2 = Picture1.Width - 90
    BottomLine.X2 = Picture1.Width - 90
    RightLine.X1 = Picture1.Width - 90
    RightLine.X2 = Picture1.Width - 90
    lf% = txtTo.Left
    txtTo.Width = Picture1.Width - 120 - lf%
    txtDate.Width = Picture1.Width - 120 - lf%
    txtcc.Width = Picture1.Width - 120 - lf%
    txtsubject.Width = Picture1.Width - 120 - lf%
    txtFrom.Width = Picture1.Width - 120 - lf%
    Picture1.Refresh
End Sub

Private Sub txtcc_KeyPress(KeyAscii As Integer)
    ' Ignora todas las pulsaciones de teclas.
    KeyAscii = 0
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    ' Ignora todas las pulsaciones de teclas.
    KeyAscii = 0
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    ' Ignora todas las pulsaciones de teclas.
    KeyAscii = 0
End Sub

Private Sub txtNoteText_KeyPress(KeyAscii As Integer)
    ' Ignora todas las pulsaciones de teclas.
    KeyAscii = 0
End Sub

Private Sub txtsubject_KeyPress(KeyAscii As Integer)
    ' Ignora todas las pulsaciones de teclas.
    KeyAscii = 0
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    ' Ignora todas las pulsaciones de teclas.
    KeyAscii = 0
End Sub

