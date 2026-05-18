VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.MDIForm VBMail 
   BackColor       =   &H8000000C&
   Caption         =   "Mail de VB"
   ClientHeight    =   4815
   ClientLeft      =   240
   ClientTop       =   810
   ClientWidth     =   9210
   Icon            =   "VBMAIL.frx":0000
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   9210
      TabIndex        =   0
      Top             =   4485
      Width           =   9204
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7260
         X2              =   7260
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   60
         X2              =   7260
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   7260
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   8580
         X2              =   8580
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7320
         X2              =   8580
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   7320
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   8580
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   10800
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label UnreadLbl 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   60
         Width           =   1575
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label TimeLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   205
         Left            =   7500
         TabIndex        =   10
         Top             =   75
         Width           =   345
      End
      Begin VB.Label MsgCountLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Count Information"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   75
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9150
      TabIndex        =   8
      Top             =   525
      Visible         =   0   'False
      Width           =   9204
      Begin VB.Timer Timer1 
         Interval        =   15000
         Left            =   180
         Top             =   120
      End
      Begin MSMAPI.MAPIMessages MapiMess 
         Left            =   1320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327680
         AddressEditFieldCount=   0
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   -1  'True
      End
      Begin MSMAPI.MAPISession MapiSess 
         Left            =   720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327680
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327680
         CancelError     =   -1  'True
         FilterIndex     =   683
         FontSize        =   1.17491e-38
      End
      Begin VB.Label Label1 
         Caption         =   "Estos controles no son visibles en tiempo de ejecución"
         Height          =   315
         Left            =   2700
         TabIndex        =   9
         Top             =   300
         Width           =   3915
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9210
      TabIndex        =   12
      Top             =   0
      Width           =   9204
      Begin VB.CommandButton Delete 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5640
         TabIndex        =   4
         Top             =   90
         Width           =   1155
      End
      Begin VB.CommandButton Next 
         Caption         =   "&Siguiente"
         Enabled         =   0   'False
         Height          =   330
         Left            =   8040
         TabIndex        =   6
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Previous 
         Caption         =   "&Anterior"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6840
         TabIndex        =   5
         Top             =   90
         Width           =   1155
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "&Reenviar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   4140
         TabIndex        =   3
         Top             =   90
         Width           =   1395
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "Responder a to&dos"
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   2580
         TabIndex        =   2
         Top             =   90
         Width           =   1515
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "&Responder"
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   1320
         TabIndex        =   1
         Top             =   90
         Width           =   1155
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "&Redacción"
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   180
         TabIndex        =   13
         Top             =   90
         Width           =   1035
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   15
         X2              =   10800
         Y1              =   505
         Y2              =   505
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&Archivo"
      Begin VB.Menu PrintMessage 
         Caption         =   "I&mprimir mensaje"
         Enabled         =   0   'False
      End
      Begin VB.Menu PrSetup 
         Caption         =   "&Configuración de la impresora..."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edición"
      Begin VB.Menu EditDelete 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Mail 
      Caption         =   "&Correo"
      Begin VB.Menu Logon 
         Caption         =   "I&nicio de sesión"
      End
      Begin VB.Menu LogOff 
         Caption         =   "&Fin de sesión"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu rMsgList 
         Caption         =   "Actualizar la lista de mensajes"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Opciones"
      Begin VB.Menu MailOpts 
         Caption         =   "&Correo..."
      End
      Begin VB.Menu FontS 
         Caption         =   "&Fuentes"
         Begin VB.Menu FontScreen 
            Caption         =   "&Pantalla..."
         End
         Begin VB.Menu FontPrt 
            Caption         =   "&Impresora..."
         End
      End
      Begin VB.Menu DispTools 
         Caption         =   "Mostrar &herramientas"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Address 
      Caption         =   "&Dirección"
      Begin VB.Menu ShowAB 
         Caption         =   "Mostrar libreta de direcciones"
      End
   End
   Begin VB.Menu Window 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu wa 
         Caption         =   "&Cascada"
         Index           =   0
      End
      Begin VB.Menu wa 
         Caption         =   "Mosaico &horizontal"
         Index           =   1
      End
      Begin VB.Menu wa 
         Caption         =   "Mosaico &vertical"
         Index           =   2
      End
      Begin VB.Menu wa 
         Caption         =   "Ali&near iconos"
         Index           =   3
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "Ay&uda"
      Begin VB.Menu About 
         Caption         =   "&Acerca de"
      End
   End
End
Attribute VB_Name = "VBMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
    MsgBox "Aplicación de ejemplo de correo electrónico", 0, "VB Mail"
End Sub

Private Sub Delete_Click()
' Elimina un mensaje del correo.

    ' Ver todos los mensajes seleccionados que están eliminados.
    If TypeOf VBMail.ActiveForm Is MsgView Then
        Call DeleteMessage
    ElseIf TypeOf VBMail.ActiveForm Is MailLst Then
        ' Eliminar selección múltiple.
        VBMail.MapiMess.MsgIndex = MailLst.MList.ListIndex
        Call DeleteMessage
    End If

End Sub

Private Sub DispTools_Click()
    DispTools.Checked = Not DispTools.Checked
    MailLst.Tools.Visible = DispTools.Checked

    
    If MailLst.Tools.Visible Then
        Factor = 1
        ToolsSize% = -MailLst.Tools.Height
    Else
        Factor = -1
        ToolsSize% = 0
    End If

    Select Case MailLst.WindowState
        Case 0    ' Cambia el tamaño del formulario para reflejar la adición o sustracción de una barra de herramientas.
            MailLst.Height = MailLst.Height + (Factor * MailLst.Tools.Height)
        Case 2    ' Si está maximizada, ajustar el tamaño del cuadro de lista.
            MailLst.MList.Height = ScaleHeight - 90 - MailLst.MList.Top + ToolsSize%
    End Select
End Sub

Private Sub EditDelete_Click()
' Elimina los elementos de la lista.
On Error GoTo Trap
    If TypeOf VBMail.ActiveForm Is MailLst Then
        Call Delete_Click
    End If
    Exit Sub

Trap:
    ' Si ocurre un error, probablemente no hay un formulario activo.
    ' Sale del procedimiento.
    Exit Sub
End Sub

Private Sub Exit_Click()
    ' Cerrar la aplicación y desconectar.
    If MapiSess.SessionID <> 0 Then
        Call logoff_Click
    End If
    End
End Sub

Private Sub FontPrt_Click()
    ' Establecer las fuentes de la impresora.
    On Error Resume Next
    CMDialog1.Flags = 2
    CMDialog1.FontName = Printer.FontName
    CMDialog1.FontSize = Printer.FontSize
    CMDialog1.FontBold = Printer.FontBold
    CMDialog1.FontItalic = Printer.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        Printer.FontName = CMDialog1.FontName
        Printer.FontSize = CMDialog1.FontSize
        Printer.FontBold = CMDialog1.FontBold
        Printer.FontItalic = CMDialog1.FontItalic
    End If

End Sub

Private Sub FontScreen_Click()
    ' Establecer las fuentes de la pantalla para el control activo.
    On Error Resume Next
    CMDialog1.Flags = 1
    CMDialog1.FontName = VBMail.ActiveForm.ActiveControl.FontName
    CMDialog1.FontSize = VBMail.ActiveForm.ActiveControl.FontSize
    CMDialog1.FontBold = VBMail.ActiveForm.ActiveControl.FontBold
    CMDialog1.FontItalic = VBMail.ActiveForm.ActiveControl.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        VBMail.ActiveForm.ActiveControl.FontName = CMDialog1.FontName
        VBMail.ActiveForm.ActiveControl.FontSize = CMDialog1.FontSize
        VBMail.ActiveForm.ActiveControl.FontBold = CMDialog1.FontBold
        VBMail.ActiveForm.ActiveControl.FontItalic = CMDialog1.FontItalic
    End If
End Sub

Private Sub logoff_Click()
    ' Desconectar del sistema de correo.
    Call LogOffUser
End Sub

Private Sub Logon_Click()
    ' Conectar al sistema de correo.
    On Error Resume Next
    MapiSess.Action = 1
    If Err <> 0 Then
        MsgBox "Error de inicio de sesión: " + Error$
    Else
        Screen.MousePointer = 11
        MapiMess.SessionID = MapiSess.SessionID
        ' Adquirir el contador de mensajes.
        GetMessageCount
        ' Cargar la lista de correo con la información de los sobres.
        Screen.MousePointer = 11
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
        ' Ajustar los botones según se necesite.
        Logon.Enabled = False
        LogOff.Enabled = True
        VBMail.SendCtl(vbMessageCompose).Enabled = True
        VBMail.SendCtl(vbMessageReplyAll).Enabled = True
        VBMail.SendCtl(vbMessageReply).Enabled = True
        VBMail.SendCtl(vbMessageForward).Enabled = True
        VBMail.PrintMessage.Enabled = True
        VBMail.DispTools.Enabled = True
        VBMail.rMsgList.Enabled = True
        VBMail.EditDelete.Enabled = True
      End If
End Sub

Private Sub MailOpts_Click()
    ' Mostrar el formulario de "Opciones de correo".
    OptionType = conOptionGeneral
    MailOptFrm.Show 1
End Sub

Private Sub MDIForm_Load()
    ' Asegurar que todos los controles cambian de tamaño cuando se necesite.
    TimeLbl = Time$
     SendWithMapi = True
     Call Picture1_Resize
     Call Picture2_Resize
     VBMail.MsgCountLbl = "Desconectado"
End Sub

Private Sub Next_Click()
    ' Ver el próximo mensaje en la lista.
    If MailLst.MList.ListIndex <> MailLst.MList.ListCount - 1 Then
        MailLst.MList.ItemData(MailLst.MList.ListIndex) = False
        MailLst.MList.ListIndex = MailLst.MList.ListIndex + 1
    End If
    Call ViewNextMsg
End Sub

Private Sub Picture1_Resize()
Const TimeBoxStartOffset = 1200
Const TimeBoxEndOffset = 60
Const MsgBoxStartOffset = 60
Const MsgBoxEndOffset = TimeBoxStartOffset + 90

    ' Ajusta los tamaños de las líneas y sitúa la etiqueta de la hora.
    TimeLbl.Left = Picture1.Width - TimeLbl.Width - 265
    TopLine2(0).X2 = Picture1.Width
    TopLine2(1).X2 = Picture1.Width

    TimeBoxLine(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(0).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxLine(1).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(1).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxSide(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxSide(0).X2 = Picture1.Width - TimeBoxStartOffset

    TimeBoxSide(1).X1 = Picture1.Width - TimeBoxEndOffset
    TimeBoxSide(1).X2 = Picture1.Width - TimeBoxEndOffset

    MsgBoxLine(0).X2 = Picture1.Width - MsgBoxEndOffset
    MsgBoxLine(1).X2 = Picture1.Width - MsgBoxEndOffset

    MsgBoxSide(1).X1 = Picture1.Width - MsgBoxEndOffset
    MsgBoxSide(1).X2 = Picture1.Width - MsgBoxEndOffset

    Picture1.Refresh
End Sub

Private Sub Picture2_Resize()
    ' Ajustar la posición de las líneas.
    TopLine(0).X2 = Picture2.Width
    TopLine(1).X2 = Picture2.Width
    Picture2.Refresh
End Sub

Private Sub Previous_Click()
    ' Ver el mensaje anterior en la lista.
    If MailLst.MList.ListIndex <> 0 Then
        MailLst.MList.ItemData(MailLst.MList.ListIndex) = False
        MailLst.MList.ListIndex = MailLst.MList.ListIndex - 1
    End If
    Call ViewNextMsg
End Sub

Private Sub PrintMessage_Click()
    ' Imprime el correo.
    Call PrintMail
End Sub

Private Sub PrSetup_Click()
' Llama al procedimiento de configuración de impresora en el control de diálogo común.
On Error Resume Next
    CMDialog1.Flags = &H40  ' Sólo el cuadro de diálogo de configuración de la impresora.
    CMDialog1.ShowPrinter
End Sub

Private Sub rMsgList_Click()
        Screen.MousePointer = 11
        GetMessageCount
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
End Sub

Private Sub SendCtl_Click(Index As Integer)
Dim NewMessage As New NewMsg
    On Error Resume Next

    ' Index = 6: Nuevo mensaje
    '       = 7: Responder
    '       = 8: Responder a todos
    '       = 9: Hacia abajo

    ' Guardar la información de cabecera y el texto de la nota actual.
    If Index > 6 Then
        ' SVNote = GetHeader(VBMAIL.MapiMess) + VBMAIL.MapiMess.MsgNoteText
        SVNote = VBMail.MapiMess.MsgNoteText
        SVNote = GetHeader(VBMail.MapiMess) + SVNote
    End If

    VBMail.MapiMess.Action = Index

    ' Establecer el texto del nuevo mensaje.
    If Index > 6 Then
        VBMail.MapiMess.MsgNoteText = SVNote
    End If

    If SendWithMapi Then
        VBMail.MapiMess.Action = vbMessageSendDlg
    Else
        Call LoadMessage(-1, NewMessage)   ' Cargar el mensaje en la ventana NewMSG de VBMail.
    End If
End Sub

Private Sub ShowAB_Click()
On Error Resume Next
    ' Muestra la dirección para el mensaje actual.
    VBMail.MapiMess.Action = vbMessageShowAdBook
    If Err Then
        If Err <> 32001 Then   ' El usuario elige "Cancelar".
            MsgBox "Error: Ocurrió un error '" + Error$ + "' al intentar mostrar la Libreta de direcciones"
        End If
    Else
        If TypeOf VBMail.ActiveForm Is NewMsg Then
            Call UpdateRecips(VBMail.ActiveForm)
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    ' Actualiza la etiqueta de la hora.
    TimeLbl = Time$
End Sub

Private Sub wa_Click(Index As Integer)
    ' Organiza las ventanas según sean seleccionadas.
    VBMail.Arrange Index
End Sub

