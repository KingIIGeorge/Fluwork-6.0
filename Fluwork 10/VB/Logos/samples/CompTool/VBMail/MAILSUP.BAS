Attribute VB_Name = "Module1"
Public Const conMailLongDate = 0
Public Const conMailListView = 1

Public Const conOptionGeneral = 1       ' Constante para el tipo de diálogo de opción - Opciones generales
Public Const conOptionMessage = 2       ' Constante para el tipo de diálogo de opción - Opciones de mensajes

Public Const conUnreadMessage = "*"     ' Constante para la cadena que indica mensaje no leído

Public Const vbRecipTypeTo = 1
Public Const vbRecipTypeCc = 2

Public Const vbMessageFetch = 1
Public Const vbMessageSendDlg = 2
Public Const vbMessageSend = 3
Public Const vbMessageSaveMsg = 4
Public Const vbMessageCopy = 5
Public Const vbMessageCompose = 6
Public Const vbMessageReply = 7
Public Const vbMessageReplyAll = 8
Public Const vbMessageForward = 9
Public Const vbMessageDelete = 10
Public Const vbMessageShowAdBook = 11
Public Const vbMessageShowDetails = 12
Public Const vbMessageResolveName = 13
Public Const vbRecipientDelete = 14
Public Const vbAttachmentDelete = 15

Public Const vbAttachTypeData = 0
Public Const vbAttachTypeEOLE = 1
Public Const vbAttachTypeSOLE = 2

Type ListDisplay
    Name As String * 20
    Subject As String * 40
    Date As String * 20
End Type

Public currentRCIndex As Integer
Public UnRead As Integer
Public SendWithMapi As Integer
Public ReturnRequest As Integer
Public OptionType As Integer

' Funciones Windows API
#If Win32 Then
    Declare Function GetProfileString Lib "kernel32" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
#Else
    Declare Function GetProfileString% Lib "Kernel" (ByVal lpSection$, ByVal lpEntry$, ByVal lpDefault$, ByVal Buffer$, ByVal cbBuffer%)
#End If

Sub Attachments(Msg As Form)
    ' Elimina la lista actual de datos adjuntos.
    Msg.aList.Clear

    ' Si hay datos adjuntos, se cargan en el cuadro de lista.
    If VBMail.MapiMess.AttachmentCount Then
        Msg.NumAtt = VBMail.MapiMess.AttachmentCount & " archivos"
        For i% = 0 To VBMail.MapiMess.AttachmentCount - 1
            VBMail.MapiMess.AttachmentIndex = i%
            a$ = VBMail.MapiMess.AttachmentName
            Select Case VBMail.MapiMess.AttachmentType
                Case vbAttachTypeData
                    a$ = a$ + " (Archivo de datos)"
                Case vbAttachTypeEOLE
                    a$ = a$ + " (Objeto OLE incrustado)"
                Case vbAttachTypeSOLE
                    a$ = a$ + " (Objeto OLE estático)"
                Case Else
                    a$ = a$ + " (Tipo de adjunto desconocido)"
            End Select
            Msg.aList.AddItem a$
        Next i%
        
        If Not Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = True
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height + Msg.AttachWin.Height
            ' End If
        End If
    
    Else
        If Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = False
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height - Msg.AttachWin.Height
            ' End If
        End If
    End If
    Msg.Refresh
End Sub

Sub CopyNamestoMsgBuffer(Msg As Form, fResolveNames As Integer)
    Call KillRecips(VBMail.MapiMess)
    Call SetRCList(Msg.txtTo, VBMail.MapiMess, vbRecipTypeTo, fResolveNames)
    Call SetRCList(Msg.txtcc, VBMail.MapiMess, vbRecipTypeCc, fResolveNames)
End Sub

Function DateFromMapiDate$(ByVal S$, wFormat%)
' Este procedimiento da formato a una fecha MAPI con uno de
' los dos formatos que existen para ver el mensaje.
    Y$ = Left$(S$, 4)
    M$ = Mid$(S$, 6, 2)
    D$ = Mid$(S$, 9, 2)
    T$ = Mid$(S$, 12)
    Ds# = DateValue(M$ + "/" + D$ + "/" + Y$) + TimeValue(T$)
    Select Case wFormat
        Case conMailLongDate
            f$ = "dddd, mmmm d, yyyy, h:mmAM/PM"
        Case conMailListView
            f$ = "mm/dd/yy hh:mm"
    End Select
    DateFromMapiDate = Format$(Ds#, f$)
End Function

Sub DeleteMessage()
    ' Si el formulario activo actualmente es un mensaje, establecer
    ' MListIndex con el valor correcto.
    If TypeOf Screen.ActiveForm Is MsgView Then
        MailLst.MList.ListIndex = Val(Screen.ActiveForm.Tag)
        ViewingMsg = True
    End If

   ' Elimina el mensaje de correo.
    If MailLst.MList.ListIndex <> -1 Then
        VBMail.MapiMess.MsgIndex = MailLst.MList.ListIndex
        VBMail.MapiMess.Action = vbMessageDelete
        X% = MailLst.MList.ListIndex
        MailLst.MList.RemoveItem X%
        If X% < MailLst.MList.ListCount - 1 Then
            MailLst.MList.ListIndex = X%
        Else
            MailLst.MList.ListIndex = MailLst.MList.ListCount - 1
        End If
        VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " mensajes"

        ' Ajusta los valores de índice para los mensajes que se están viendo
        ' actualmente.
        If ViewingMsg Then
            Screen.ActiveForm.Tag = Str$(-1)
        End If

        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) > X% Then
                    Forms(i).Tag = Val(Forms(i).Tag) - 1
                End If
            End If
        Next i
        
        ' Si el usuario está viendo un mensaje, carga el próximo mensaje
        ' en el formulario MsgView, si el mensaje no se está mostrando
        ' actualmente.
        If ViewingMsg Then
            ' Primero verifica si el mensaje se está viendo actualmente.
            WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
            If WindowNum% > 0 Then
                If Forms(WindowNum%).Caption <> Screen.ActiveForm.Caption Then
                    Unload Screen.ActiveForm
                     ' Se encuentra la ventana correcta de nuevo y se muestra.
                     ' El índice no es valido después de liberar.
                     Forms(FindMsgWindow((MailLst.MList.ListIndex))).Show
                Else
                     Forms(WindowNum%).Show
                End If
            Else
                Call LoadMessage(MailLst.MList.ListIndex, Screen.ActiveForm)
            End If
        Else
            ' Verifica si había una ventana mostrando el mensaje y libera la pantalla.
            WindowNum% = FindMsgWindow(X%)
            If WindowNum% > 0 Then
                Unload Forms(X%)
            End If
        End If
     End If
End Sub

Sub DisplayAttachedFile(ByVal FileName As String)
On Error Resume Next
        ' Determina la extensión del archivo.
        ext$ = FileName
        junk$ = Token$(ext$, ".")
        ' Abre la aplicación desde el archivo WIN.INI.
        Buffer$ = String$(256, " ")
        errCode% = GetProfileString("Extensions", ext$, "NOTFOUND", Buffer$, Len(Left(Buffer$, Chr(0)) - 1))
        If errCode% Then
            Buffer$ = Mid$(Buffer$, 1, InStr(Buffer$, Chr(0)) - 1)
            If Buffer$ <> "NOTFOUND" Then
                ' Quitar la información ^.EXT de la cadena.
                EXEName$ = Token$(Buffer$, " ")
                errCode% = Shell(EXEName$ + " " + FileName, 1)
                If Err Then
                    MsgBox "Ocurrió un error: " + Error$
                End If
            Else
                MsgBox "Aplicación que utiliza: no se encontró <" + ext$ + "> en WIN.INI"
            End If
        End If
End Sub

Function FindMsgWindow(Index As Integer) As Integer
' Esta función busca en las ventanas activas,
' localizando aquellas con el tipo MsgView y
' verificando si la etiqueta contiene el índice que el
' usuario está buscando.
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) = Index Then
                    FindMsgWindow = i
                    Exit Function
                End If
            End If
        Next i
        FindMsgWindow = -1
End Function

Function GetHeader(Msg As Control) As String
Dim CR As String
CR = Chr$(13) + Chr$(10)
      Header$ = String$(25, "-") + CR
      Header$ = Header$ + "De: " + Msg.MsgOrigDisplayName + CR
      Header$ = Header$ + "Para: " + GetRCList(Msg, vbRecipTypeTo) + CR
      Header$ = Header$ + "Cc: " + GetRCList(Msg, vbRecipTypeCc) + CR
      Header$ = Header$ + "Asunto: " + Msg.MsgSubject + CR
      Header$ = Header$ + "Fecha: " + DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate) + CR + CR
      GetHeader = Header$
End Function

Sub GetMessageCount()
    '  Lee todos los mensajes de correo y muestra el contador.
    Screen.MousePointer = 11
    VBMail.MapiMess.FetchUnreadOnly = 0
    VBMail.MapiMess.Action = vbMessageFetch
    VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " mensajes"
    Screen.MousePointer = 0
End Sub

Function GetRCList(Msg As Control, RCType As Integer) As String
' Dada una lista de receptores, esta función devuelve
' una lista de receptores del tipo especificado con el formato siguiente:
'
'       Persona 1;Persona 2;Persona 3

    For i = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = i
        If RCType = Msg.RecipType Then
                a$ = a$ + ";" + Msg.RecipDisplayName
        End If
    Next i
    If a$ <> "" Then
       a$ = Mid$(a$, 2)  ' Quita el ";" precedente.
    End If
    GetRCList = a$
End Function

Sub KillRecips(MsgControl As Control)
    ' Borrar cada receptor.  Repetir hasta que no queden receptores.
    While MsgControl.RecipCount
        MsgControl.Action = vbRecipientDelete
    Wend
End Sub

Sub LoadList(mailctl As Control)
' Este procedimiento carga las cabeceras de los mensajes del correo
' en la lista MailLst.MList. Los mensajes que no se han leído tienen
' conUnreadMessage al principio de la cadena.
    MailLst.MList.Clear
    UnRead = 0
    StartIndex = 0
    For i = 0 To mailctl.MsgCount - 1
        mailctl.MsgIndex = i
        If Not mailctl.MsgRead Then
            a$ = conUnreadMessage + " "
            If UnRead = 0 Then
                StartIndex = i  ' Posición de inicio en la lista de correo.
            End If
            UnRead = UnRead + 1
        Else
            a$ = "  "
        End If
        a$ = a$ + Mid$(Format$(mailctl.MsgOrigDisplayName, "!" + String$(10, "@")), 1, 10)
        If mailctl.MsgSubject <> "" Then
            b$ = Mid$(Format$(mailctl.MsgSubject, "!" + String$(35, "@")), 1, 35)
        Else
            b$ = String$(30, " ")
        End If
        c$ = Mid$(Format$(DateFromMapiDate(mailctl.MsgDateReceived, conMailListView), "!" + String$(15, "@")), 1, 15)
        MailLst.MList.AddItem a$ + Chr$(9) + b$ + Chr$(9) + c$
        MailLst.MList.Refresh
    Next i

    MailLst.MList.ListIndex = StartIndex
    
    ' Habilita los botones correctos.
    VBMail.Next.Enabled = True
    VBMail.Previous.Enabled = True
    VBMail![Delete].Enabled = True

    ' Ajustar el valor de las etiquetas mostrando el mensaje del contador.
    If UnRead Then
        VBMail.UnreadLbl = " - " + Format$(UnRead) + " sin leer"
        MailLst.Icon = MailLst.NewMail.Picture
    Else
        VBMail.UnreadLbl = ""
        MailLst.Icon = MailLst.nonew.Picture
    End If
End Sub
    

Sub LoadMessage(ByVal Index As Integer, Msg As Form)
' Este procedimiento carga el mensaje de correo especificado en
' un formulario para ver o editar un mensaje.
    If TypeOf Msg Is MsgView Then
        a$ = MailLst.MList.List(Index)
        ' El mensaje no se ha leído; Se restablece el texto.
        If Mid$(a$, 1, 1) = conUnreadMessage Then
            Mid$(a$, 1, 1) = " "
            MailLst.MList.List(Index) = a$
            UnRead = UnRead - 1
            If UnRead Then
                VBMail.UnreadLbl = Format$(UnRead) + " sin leer"
            Else
                VBMail.UnreadLbl = ""
                ' Cambia el icono en la ventana de lista.
                MailLst.Icon = MailLst.nonew.Picture
            End If
        End If
    End If

    ' Estos campos sólo afectan al formulario de ver en pantalla.
    If TypeOf Msg Is MsgView Then
        VBMail.MapiMess.MsgIndex = Index
        Msg.txtDate = DateFromMapiDate$(VBMail.MapiMess.MsgDateReceived, conMailLongDate)
        Msg.txtFrom = VBMail.MapiMess.MsgOrigDisplayName
        MailLst.MList.ItemData(Index) = True
    End If
    ' Estos campos afectan a los dos tipos de formulario.
    Call Attachments(Msg)
    Msg.txtNoteText = VBMail.MapiMess.MsgNoteText
    Msg.txtsubject = VBMail.MapiMess.MsgSubject
    Msg.Caption = VBMail.MapiMess.MsgSubject
    Msg.Tag = Index
    Call UpdateRecips(Msg)
    Msg.Refresh
    Msg.Show
End Sub

Sub LogOffUser()
    On Error Resume Next
    VBMail.MapiSess.Action = 2
    If Err <> 0 Then
        MsgBox "Error en la desconexión: " + ErrorR
    Else
        VBMail.MapiMess.SessionID = 0
        ' Ajusta los elementos del menú.
        VBMail.LogOff.Enabled = 0
        VBMail.Logon.Enabled = -1
        ' Libera todos los formularios excepto el formulario MDI.
        Do Until Forms.Count = 1
            i = Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                ' No hacer nada.
            Else
                Unload Forms(i)
            End If
        Loop
        ' Deshabilita los botones de la barra de herramientas.
        VBMail.Next.Enabled = False
        VBMail.Previous.Enabled = False
        VBMail![Delete].Enabled = False
        VBMail.SendCtl(vbMessageCompose).Enabled = False
        VBMail.SendCtl(vbMessageReplyAll).Enabled = False
        VBMail.SendCtl(vbMessageReply).Enabled = False
        VBMail.SendCtl(vbMessageForward).Enabled = False
        VBMail.rMsgList.Enabled = False
        VBMail.PrintMessage.Enabled = False
        VBMail.DispTools.Enabled = False
        VBMail.EditDelete.Enabled = False
                          
        ' Restablece el título para las etiquetas de la barra de estado.
        VBMail.MsgCountLbl = "Desconectado"
        VBMail.UnreadLbl = ""
    End If

End Sub

Sub PrintLongText(ByVal LongText As String)
' Este procedimiento imprime un flujo de texto en una impresora y
' se asegura que las palabras no están cortadas entre líneas y
' que se ajustan según sea necesario.
    Do Until LongText = ""
        Word$ = Token$(LongText, " ")
        If Printer.TextWidth(Word$) + Printer.CurrentX > Printer.Width - Printer.TextWidth("ZZZZZZZZ") Then
            Printer.Print
        End If
        Printer.Print " " + Word$;
    Loop
End Sub

Sub PrintMail()
    ' En modo de ver lista, todos los mensajes seleccionados se imprimen.
    ' En modo de ver mensaje, el mensaje seleccionado se imprime.

    If TypeOf Screen.ActiveForm Is MsgView Then
        Call PrintMessage(VBMail.MapiMess, False)
        Printer.EndDoc
    ElseIf TypeOf Screen.ActiveForm Is MailLst Then
        For i = 0 To MailLst.MList.ListCount - 1
            If MailLst.MList.Selected(i) Then
                VBMail.MapiMess.MsgIndex = i
                Call PrintMessage(VBMail.MapiMess, False)
            End If
        Next i
        Printer.EndDoc
    End If
End Sub

Sub PrintMessage(Msg As Control, fNewPage As Integer)
'   Este procedimiento imprime un mensaje de correo.
    Screen.MousePointer = 11
    ' Comienza una nueva página si es necesario.
    If fNewPage Then
        Printer.NewPage
    End If
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.DrawWidth = 10
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.Print
    Printer.FontSize = 9.75
    Printer.Print "De:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgOrigDisplayName
    Printer.Print "Para:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, vbRecipTypeTo)
    Printer.Print "Cc:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, vbRecipTypeCc)
    Printer.Print "Asunto:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgSubject
    Printer.Print "Fecha:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate)
    Printer.Print
    Printer.DrawWidth = 5
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.FontSize = 9.75
    Printer.FontBold = False
    Call PrintLongText(Msg.MsgNoteText)
    Printer.Print
    Screen.MousePointer = 0
End Sub

Sub SaveMessage(Msg As Form)
    ' Guarda el asunto y el texto de la nota.
    ' Copia el mensaje al búfer.
    ' Restablece el asunto y el texto del mensaje.
    ' Guarda el mensaje.
    svSub = Msg.txtsubject
    SVNote = Msg.txtNoteText
    VBMail.MapiMess.Action = vbMessageCopy
    VBMail.MapiMess.MsgSubject = svSub
    VBMail.MapiMess.MsgNoteText = SVNote
    VBMail.MapiMess.Action = vbMessageSaveMsg
End Sub

Sub SetRCList(ByVal NameList As String, Msg As Control, RCType As Integer, fResolveNames As Integer)
' Dada una lista de receptores:
'
'       Persona 1;Persona 2;Persona 3
'
' este procedimiento la coloca en las estructuras de Msg.Recip.
    
    If NameList = "" Then
        Exit Sub
    End If

    i = Msg.RecipCount
    Do
        Msg.RecipIndex = i
        Msg.RecipDisplayName = Trim$(Token(NameList, ";"))
        If fResolveNames Then
            Msg.Action = vbMessageResolveName
        End If
        Msg.RecipType = RCType
        i = i + 1
    Loop Until (NameList = "")
End Sub

Sub SizeMessageWindow(MsgWindow As Form)
    If MsgWindow.WindowState <> 1 Then
        ' Determinar el tamaño mínimo de la ventana basándose
        ' en la visibilidad de AttachWin (Ventana de datos adjuntos).
        If MsgWindow.AttachWin.Visible Then    ' Ventana de datos adjuntos.
            MinSize = 3700
        Else
            MinSize = 3700 - MsgWindow.AttachWin.Height
        End If

        ' Mantiene el tamaño mínimo del formulario.
        If MsgWindow.Height < MinSize And (MsgWindow.WindowState = 0) Then
            MsgWindow.Height = MinSize
            Exit Sub

        End If
        ' Ajusta el tamaño del cuadro de texto.
        If MsgWindow.ScaleHeight > MsgWindow.txtNoteText.Top Then
            If MsgWindow.AttachWin.Visible Then
                X% = MsgWindow.AttachWin.Height
            Else
                X% = 0
            End If
            MsgWindow.txtNoteText.Height = MsgWindow.ScaleHeight - MsgWindow.txtNoteText.Top - X%
            MsgWindow.txtNoteText.Width = MsgWindow.ScaleWidth
        End If
    End If

End Sub

Function Token$(tmp$, search$)
    X = InStr(1, tmp$, search$)
    If X Then
       Token$ = Mid$(tmp$, 1, X - 1)
       tmp$ = Mid$(tmp$, X + 1)
    Else
       Token$ = tmp$
       tmp$ = ""
    End If
End Function

Sub UpdateRecips(Msg As Form)
' Este procedimiento actualiza los campos correctos de edición
' y la información de los receptores.
    Msg.txtTo.Text = GetRCList(VBMail.MapiMess, vbRecipTypeTo)
    Msg.txtcc.Text = GetRCList(VBMail.MapiMess, vbRecipTypeCc)
End Sub

Sub ViewNextMsg()
    ' Verificar si el mensaje está cargado actualmente.
    ' Si está cargado, muestra este formulario.
    ' Si no está cargado, carga el mensaje.
    WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
    If WindowNum% > 0 Then
        Forms(WindowNum%).Show
    Else
        If TypeOf Screen.ActiveForm Is MsgView Then
            Call LoadMessage(MailLst.MList.ListIndex, Screen.ActiveForm)
        Else
            Dim Msg As New MsgView
            Call LoadMessage(MailLst.MList.ListIndex, Msg)
        End If
    End If
End Sub

