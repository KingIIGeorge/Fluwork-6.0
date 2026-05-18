VERSION 5.00
Begin VB.Form NewMsg 
   Caption         =   "Enviar nota"
   ClientHeight    =   3135
   ClientLeft      =   780
   ClientTop       =   1245
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   8070
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8070
      TabIndex        =   8
      Top             =   0
      Width           =   8070
      Begin VB.CommandButton CompAdd 
         Caption         =   "Dir&ección"
         Height          =   330
         Left            =   6480
         TabIndex        =   13
         Top             =   90
         Width           =   1455
      End
      Begin VB.CommandButton CompOpt 
         Caption         =   "&Opciones"
         Height          =   330
         Left            =   4920
         TabIndex        =   12
         Top             =   90
         Width           =   1455
      End
      Begin VB.CommandButton Attach 
         Caption         =   "A&djuntar"
         Height          =   330
         Left            =   3360
         TabIndex        =   11
         Top             =   90
         Width           =   1455
      End
      Begin VB.CommandButton ChkNames 
         Caption         =   "Co&mprobar nombres"
         Height          =   330
         Left            =   1680
         TabIndex        =   10
         Top             =   90
         Width           =   1575
      End
      Begin VB.CommandButton Send 
         Caption         =   "En&viar"
         Height          =   330
         Left            =   180
         TabIndex        =   9
         Top             =   90
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7380
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.TextBox txtNoteText 
      Height          =   1275
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   7755
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   8070
      TabIndex        =   4
      Top             =   495
      Width           =   8070
      Begin VB.TextBox txtTo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   0
         Top             =   180
         Width           =   6915
      End
      Begin VB.TextBox txtcc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   540
         Width           =   6915
      End
      Begin VB.TextBox txtsubject 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   900
         Width           =   6915
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   7920
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1380
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   7920
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Para:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cc:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Asunto:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   720
      End
   End
End
Attribute VB_Name = "NewMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Attach_Click()
' Controla los datos adjuntos.
On Error Resume Next
   VBMail.CMDialog1.DialogTitle = "Adjuntar"
   VBMail.CMDialog1.Filter = "Todos los archivos(*.*)|*.*|Archivos de texto(*.txt)|*.txt"
   VBMail.CMDialog1.ShowOpen
   If Err = 0 Then
        On Error GoTo 0
        VBMail.MapiMess.AttachmentIndex = VBMail.MapiMess.AttachmentCount
        VBMail.MapiMess.AttachmentName = VBMail.CMDialog1.FileTitle
        VBMail.MapiMess.AttachmentPathName = VBMail.CMDialog1.FileName
        VBMail.MapiMess.AttachmentPosition = VBMail.MapiMess.AttachmentIndex
        VBMail.MapiMess.AttachmentType = vbAttachTypeData
   End If
End Sub

Private Sub ChkNames_Click()
    ' Resuelve los nombres.
    Call CopyNamestoMsgBuffer(Me, True)
    Call UpdateRecips(Me)
End Sub

Private Sub CompAdd_Click()
    ' Muestra la libreta de direcciones y la actualiza al volver.
    Call CopyNamestoMsgBuffer(Me, False)
    VBMail.MapiMess.Action = vbMessageShowAdBook
    Call UpdateRecips(Me)
End Sub

Private Sub CompOpt_Click()
    ' Muestra el formulario "Opciones del mensaje".
    OptionType = conOptionMessage
    MailOptFrm.Show 1
End Sub

Private Sub Form_Activate()
    ' Establece el MessageIndex a -1 (compone el búfer)
    ' cuando se activa esta ventana.
    VBMail.MapiMess.MsgIndex = -1
End Sub

Private Sub Form_Load()
    ' Se asegura de que las ventanas cambian al tamaño adecuado.
    Call Picture1_Resize
    Call Picture2_Resize
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    ' Ajusta el tamaño de las ventanas si el formulario no está minimizado.
    If WindowState <> 1 Then
        If ScaleHeight > txtNoteText.Top Then
            txtNoteText.Height = ScaleHeight - txtNoteText.Top
            txtNoteText.Width = ScaleWidth
        End If
    End If
End Sub

Private Sub Picture1_Resize()
    ' Actualiza el ancho de los campos y ajusta los controles
    ' de línea según sea necesario.
    TopLine(0).X2 = Picture1.Width
    TopLine(1).X2 = Picture1.Width
    Picture1.Refresh
End Sub

Private Sub Picture2_Resize()
    ' Actualiza el ancho de los campos y ajusta los controles
    ' de línea según sea necesario.
    TopLine2.X2 = Picture2.Width
    Picture2.Refresh
End Sub

Private Sub Send_Click()
    ' Coloca el texto de Subject y Note en el búfer.
    ' Agrega sitio al principio para adjuntar archivos.
    If VBMail.MapiMess.AttachmentCount > 0 Then
        txtNoteText = String$(VBMail.MapiMess.AttachmentCount, "*") + txtNoteText
    End If
    VBMail.MapiMess.MsgSubject = txtsubject
    VBMail.MapiMess.MsgNoteText = txtNoteText
    VBMail.MapiMess.MsgReceiptRequested = ReturnRequest
    Call CopyNamestoMsgBuffer(Me, True)
                  
    On Error Resume Next
    VBMail.MapiMess.Action = vbMessageSend
    If Err Then
        MsgBox "Error durante el envío: " + Str$(Err)
    Else
        Unload Me
    End If
End Sub

