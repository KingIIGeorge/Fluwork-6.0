VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmTerminal 
   Caption         =   "Terminal de Visual Basic"
   ClientHeight    =   4935
   ClientLeft      =   2940
   ClientTop       =   2055
   ClientWidth     =   7155
   ForeColor       =   &H00000000&
   Icon            =   "vbterm.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7155
   Begin ComctlLib.Toolbar tbrToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OpenLogFile"
            Description     =   "Open Log File..."
            Object.ToolTipText     =   "Abrir archivo de registro..."
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "CloseLogFile"
            Description     =   "Close Log File"
            Object.ToolTipText     =   "Cerrar archivo de registro"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DialPhoneNumber"
            Description     =   "Dial Phone Number..."
            Object.ToolTipText     =   "Marcar número de teléfono"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "HangUpPhone"
            Description     =   "Hang Up Phone"
            Object.ToolTipText     =   "Colgar teléfono"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Propiedades"
            Description     =   "Properties..."
            Object.ToolTipText     =   "Propiedades..."
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "TransmitTextFile"
            Description     =   "Transmit Text File..."
            Object.ToolTipText     =   "Transmitir archivo de texto"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "vbterm.frx":030A
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   4000
         TabIndex        =   2
         Top             =   75
         Width           =   240
         Begin VB.Image imgConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":0326
            Stretch         =   -1  'True
            ToolTipText     =   "Alternar puerto"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgNotConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":0470
            Stretch         =   -1  'True
            ToolTipText     =   "Alternar puerto"
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   210
      Top             =   3645
   End
   Begin VB.TextBox txtTerm 
      Height          =   3690
      Left            =   1245
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1140
      Width           =   5790
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   165
      Top             =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   45
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin MSComDlg.CommonDialog OpenLog 
      Left            =   105
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DefaultExt      =   "LOG"
      FileName        =   "Open Communications Log File"
      Filter          =   "Log File (*.log)|*.log;"
      FilterIndex     =   501
      FontSize        =   1.17491e-38
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Estado:"
            TextSave        =   "Estado:"
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Estado del puerto de comunicaciones"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8310
            MinWidth        =   2
            Text            =   "Valores:"
            TextSave        =   "Valores:"
            Key             =   "Settings"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Valores del puerto de comunicaciones"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1244
            TextSave        =   ""
            Key             =   "ConnectTime"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Tiempo de conexión"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "vbterm.frx":05BA
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":05D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":08F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":0C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":0F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":123E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":1558
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuOpenLog 
         Caption         =   "&Abrir archivo de registro..."
      End
      Begin VB.Menu mnuCloseLog 
         Caption         =   "&Cerrar archivo de registro"
         Enabled         =   0   'False
      End
      Begin VB.Menu M3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendText 
         Caption         =   "&Transmitir archivo de texto"
         Enabled         =   0   'False
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "&Puerto"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Puerto abierto"
      End
      Begin VB.Menu MBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Pr&opiedades..."
      End
   End
   Begin VB.Menu mnuMSComm 
      Caption         =   "&MSComm"
      Begin VB.Menu mnuInputLen 
         Caption         =   "&InputLen..."
      End
      Begin VB.Menu mnuRThreshold 
         Caption         =   "&RThreshold..."
      End
      Begin VB.Menu mnuSThreshold 
         Caption         =   "&SThreshold..."
      End
      Begin VB.Menu mnuParRep 
         Caption         =   "P&arityReplace..."
      End
      Begin VB.Menu mnuDTREnable 
         Caption         =   "&DTREnable"
      End
      Begin VB.Menu Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHCD 
         Caption         =   "&CDHolding..."
      End
      Begin VB.Menu mnuHCTS 
         Caption         =   "CTSH&olding..."
      End
      Begin VB.Menu mnuHDSR 
         Caption         =   "DSRHo&lding..."
      End
   End
   Begin VB.Menu mnuCall 
      Caption         =   "&Llamar"
      Begin VB.Menu mnuDial 
         Caption         =   "&Marcar número de teléfono"
      End
      Begin VB.Menu mnuHangUp 
         Caption         =   "&Colgar teléfono"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' VBTerm - Este es un programa de demostración del
' control ActiveX de comunicaciones MSComm.
'
' Copyright (c) 1994, Crescent Software, Inc.
' por Don Malin y Carl Franklin.
'
' Actualizado por Mike Maddox
'--------------------------------------------------
Option Explicit
                        
Dim Ret As Integer      ' Entero auxiliar.
Dim Temp As String      ' Cadena auxiliar.
Dim hLogFile As Integer ' Controlador de archivo de registro abierto.
Dim StartTime As Date   ' Almacena la hora de inicio del cronómetro del puerto

Private Sub Form_Load()
    Dim CommPort As String, Handshaking As String, Settings As String
        
    On Error Resume Next
    
    ' Establece el color predeterminado del terminal
    txtTerm.SelLength = Len(txtTerm)
    txtTerm.SelText = ""
    txtTerm.ForeColor = vbBlue
       
    ' Establece el título
    App.Title = "Terminal de Visual Basic"
    
    ' Establece la luz indicadora de estado
    imgNotConnected.ZOrder
       
    ' Centra el formulario
    frmTerminal.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    ' Carga la configuración del registro
    
    Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
    If Settings <> "" Then
        MSComm1.Settings = Settings
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    CommPort = GetSetting(App.Title, "Properties", "CommPort", "") ' frmTerminal.MSComm1.CommPort
    If CommPort <> "" Then MSComm1.CommPort = CommPort
    
    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'frmTerminal.MSComm1.Handshaking
    If Handshaking <> "" Then
        MSComm1.Handshaking = Handshaking
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
    On Error GoTo 0

End Sub

Private Sub Form_Resize()
   ' Cambia el tamaño del control Term (ventana)
   txtTerm.Move 0, tbrToolBar.Height, frmTerminal.ScaleWidth, frmTerminal.ScaleHeight - sbrStatus.Height - tbrToolBar.Height
   
   ' Sitúa la luz indicadora de estado
   Frame1.Left = ScaleWidth - Frame1.Width * 1.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Counter As Long

    If MSComm1.PortOpen Then
       ' Espera 10 segundos para transmitir los datos.
       Counter = Timer + 10
       Do While MSComm1.OutBufferCount
          Ret = DoEvents()
          If Timer > Counter Then
             Select Case MsgBox("Imposible enviar los datos", 34)
                ' Cancelar.
                Case 3
                   Cancel = True
                   Exit Sub
                ' Reintentar.
                Case 4
                   Counter = Timer + 10
                ' Ignorar.
                Case 5
                   Exit Do
             End Select
          End If
       Loop

       MSComm1.PortOpen = 0
    End If

    ' Si el archivo de registro está abierto, vuelca y lo cierra.
    If hLogFile Then mnuCloseLog_Click
    End
End Sub

Private Sub imgConnected_Click()
    ' Llama a la rutina mnuOpen_Click para alternar entre conectar y desconectar.
    Call mnuOpen_Click
End Sub

Private Sub imgNotConnected_Click()
    ' Llama a la rutina mnuOpen_Click para alternar entre conectar y desconectar.
    Call mnuOpen_Click
End Sub

Private Sub mnuCloseLog_Click()
    ' Cierra el archivo de registro.
    Close hLogFile
    hLogFile = 0
    mnuOpenLog.Enabled = True
    tbrToolBar.Buttons("OpenLogFile").Enabled = True
    mnuCloseLog.Enabled = False
    tbrToolBar.Buttons("CloseLogFile").Enabled = False
    frmTerminal.Caption = "Terminal de Visual Basic"
End Sub

Private Sub mnuDial_Click()
    On Local Error Resume Next
    Static Num As String
    
    Num = "1-206-936-6735" ' Este es el número de MSDN
    
    ' Obtiene un número del usuario.
    Num = InputBox$("Escriba el número de teléfono:", "Marcar número", Num)
    If Num = "" Then Exit Sub
    
    ' Abre el puerto si no está abierto ya.
    If Not MSComm1.PortOpen Then
       mnuOpen_Click
       If Err Then Exit Sub
    End If
      
    ' Habilita el botón de colgar y el elemento de menú correspondiente.
    mnuHangUp.Enabled = True
    tbrToolBar.Buttons("HangUpPhone").Enabled = True
              
    ' Marca el número.
    MSComm1.Output = "ATDT" & Num & vbCrLf
    
    ' Inicia el cronómetro del puerto.
    StartTiming
End Sub

' Alterna la propiedad DTREnabled.
Private Sub mnuDTREnable_Click()
    ' Alterna la propiedad DTREnable
    MSComm1.DTREnable = Not MSComm1.DTREnable
    mnuDTREnable.Checked = MSComm1.DTREnable
End Sub


Private Sub mnuFileExit_Click()
    ' Utiliza Form_Unload, ya que tiene código para comprobar si hay datos
    ' no enviados y si hay un archivo de registro abierto.
    Form_Unload Ret
End Sub



' Alterna la propiedad DTREnable para colgar la línea.
Private Sub mnuHangup_Click()
    On Error Resume Next
    
    MSComm1.Output = "ATH"      ' Envía la cadena de colgar
    Ret = MSComm1.DTREnable     ' Guarda el valor actual.
    MSComm1.DTREnable = True    ' Activa DTR.
    MSComm1.DTREnable = False   ' Desactiva DTR.
    MSComm1.DTREnable = Ret     ' Restablece el valor anterior.
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    
    ' Si el puerto continúa abierto, lo cierra
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
    ' Notifica el error al usuario
    If Err Then MsgBox Error$, 48
    
    mnuSendText.Enabled = False
    tbrToolBar.Buttons("TransmitTextFile").Enabled = False
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    mnuDial.Enabled = True
    tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
    sbrStatus.Panels("Settings").Text = "Valores:"
    
    ' Apaga la luz indicadora y quita la marca al menú
    mnuOpen.Checked = False
    imgNotConnected.ZOrder
            
    ' Detiene el cronómetro del puerto
    StopTiming
    sbrStatus.Panels("Status").Text = "Estado:"
    On Error GoTo 0
End Sub

' Muestra el valor de la propiedad CDHolding.
Private Sub mnuHCD_Click()
    If MSComm1.CDHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CDHolding = " + Temp
End Sub

' Muestra el valor de la propiedad CTSHolding.
Private Sub mnuHCTS_Click()
    If MSComm1.CTSHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CTSHolding = " + Temp
End Sub

' Muestra el valor de la propiedad DSRHolding.
Private Sub mnuHDSR_Click()
    If MSComm1.DSRHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "DSRHolding = " + Temp
End Sub

' Este procedimiento establece la propiedad InputLen, que determina
' el número de bytes de datos leídos cada vez que se usa Input para
' recuperar datos del búfer de entrada.
' Al establecer 0 en InputLen se especifica que debe leerse todo el
' contenido del búfer.
Private Sub mnuInputLen_Click()
    On Error Resume Next

    Temp = InputBox$("Introduzca nuevo InputLen:", "InputLen", Str$(MSComm1.InputLen))
    If Len(Temp) Then
        MSComm1.InputLen = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

Private Sub mnuProperties_Click()
  ' Muestra el formulario de propiedades de CommPort
  frmProperties.Show vbModal
  
End Sub

' Alterna el estado del puerto (abierto o cerrado).
Private Sub mnuOpen_Click()
    On Error Resume Next
    Dim OpenFlag

    MSComm1.PortOpen = Not MSComm1.PortOpen
    If Err Then MsgBox Error$, 48
    
    OpenFlag = MSComm1.PortOpen
    
    mnuOpen.Checked = OpenFlag
    mnuSendText.Enabled = OpenFlag
    tbrToolBar.Buttons("TransmitTextFile").Enabled = OpenFlag
        
    If MSComm1.PortOpen Then
        ' Habilita el botón de marcar y el elemento de menú asociado
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' Habilita el botón de colgar y el elemento de menú asociado
        mnuHangUp.Enabled = True
        tbrToolBar.Buttons("HangUpPhone").Enabled = True
        
        imgConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Valores:" & MSComm1.Settings
        StartTiming
    Else
        ' Habilita el botón de marcar y el elemento de menú asociado
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' deshabilita el botón de colgar y el elemento de menú asociado
        mnuHangUp.Enabled = False
        tbrToolBar.Buttons("HangUpPhone").Enabled = False
        
        imgNotConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Valores:"
        StopTiming
    End If
    
End Sub

Private Sub mnuOpenLog_Click()
   Dim replace
   On Error Resume Next
   OpenLog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
   OpenLog.CancelError = True
      
   ' Obtiene del usuario el nombre de archivo largo.
   OpenLog.DialogTitle = "Abrir archivo de registro de comunicaciones"
   OpenLog.Filter = "Archivos de registro (*.LOG)|*.log|Todos los archivos (*.*)|*.*"
   
   Do
      OpenLog.filename = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then Exit Sub
      Temp = OpenLog.filename

      ' Si el archivo ya existe, pregunta al usuario si desea
      ' sobrescribirlo o agregarle datos.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
      If Ret Then
         replace = MsgBox("Reemplazar el archivo - " + Temp + "?", 35)
      Else
         replace = 0
      End If
   Loop While replace = 2

   ' El usuario ha hecho clic en el botón Sí, así que elimina el archivo.
   If replace = 6 Then
      Kill Temp
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
   End If

   ' Abre el archivo de registro.
   hLogFile = FreeFile
   Open Temp For Binary Access Write As hLogFile
   If Err Then
      MsgBox Error$, 48
      Close hLogFile
      hLogFile = 0
      Exit Sub
   Else
      ' Va al final del archivo para poder agregarle nuevos datos.
      Seek hLogFile, LOF(hLogFile) + 1
   End If

   frmTerminal.Caption = "Terminal de Visual Basic - " + OpenLog.FileTitle
   mnuOpenLog.Enabled = False
   tbrToolBar.Buttons("OpenLogFile").Enabled = False
   mnuCloseLog.Enabled = True
   tbrToolBar.Buttons("CloseLogFile").Enabled = True
End Sub

' Este procedimiento establece la propiedad ParityReplace, que
' contiene el carácter que reemplazará a los caracteres
' incorrectos recibidos a causa de un error de paridad.
Private Sub mnuParRep_Click()
    On Error Resume Next

    Temp = InputBox$("Escriba el carácter a reemplazar", "ParityReplace", frmTerminal.MSComm1.ParityReplace)
    frmTerminal.MSComm1.ParityReplace = Left$(Temp, 1)
    If Err Then MsgBox Error$, 48
End Sub

' Este procedimiento establece la propiedad RThreshlold, que
' determina el número de bytes que pueden llegar al búfer de
' recepción antes de disparar el evento OnComm y de que se
' establezca comEvReceive en la propiedad CommEvent.
Private Sub mnuRThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Introduzca el nuevo RThreshold:", "RThreshold", Str$(MSComm1.RThreshold))
    If Len(Temp) Then
        MSComm1.RThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If

End Sub

' El evento OnComm se usa para interceptar eventos y errores de comunicaciones.
Private Static Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    
    ' Bifurca según la propiedad CommEvent.
    Select Case MSComm1.CommEvent
        ' Mensajes de evento.
        Case comEvReceive
            Dim Buffer As Variant
            Buffer = MSComm1.Input
            Debug.Print "Recibir - " & StrConv(Buffer, vbUnicode)
            ShowData txtTerm, (StrConv(Buffer, vbUnicode))
        Case comEvSend
        Case comEvCTS
            EVMsg$ = "Detectado cambio en CTS"
        Case comEvDSR
            EVMsg$ = "Detectado cambio en DSR"
        Case comEvCD
            EVMsg$ = "Detectado cambio en CD"
        Case comEvRing
            EVMsg$ = "El teléfono está sonando"
        Case comEvEOF
            EVMsg$ = "Detectado el final del archivo"

        ' Mensajes de error.
        Case comBreak
            ERMsg$ = "Parada recibida"
        Case comCDTO
            ERMsg$ = "Sobrepasado el tiempo de espera de detección de portadora"
        Case comCTSTO
            ERMsg$ = "Soprepasado el tiempo de espera de CTS"
        Case comDCB
            ERMsg$ = "Error recibiendo DCB"
        Case comDSRTO
            ERMsg$ = "Sobrepasado el tiempo de espera de DSR"
        Case comFrame
            ERMsg$ = "Error de marco"
        Case comOverrun
            ERMsg$ = "Error de sobrecarga"
        Case comRxOver
            ERMsg$ = "Desbordamiento en el búfer de recepción"
        Case comRxParity
            ERMsg$ = "Error de paridad"
        Case comTxFull
            ERMsg$ = "Búfer de transmisión lleno"
        Case Else
            ERMsg$ = "Error o evento desconocido"
    End Select
    
    If Len(EVMsg$) Then
        ' Muestra los mensajes de evento en la barra de estado.
        sbrStatus.Panels("Status").Text = "Estado:" & EVMsg$
                
        ' Activa el cronómetro para que el mensaje de la barra
        ' de estado se borre después de dos segundos.
        Timer2.Enabled = True
        
    ElseIf Len(ERMsg$) Then
        ' Muestra los mensajes de evento en la barra de estado.
        sbrStatus.Panels("Status").Text = "Estado:" & ERMsg$
        
        ' Muestra los mensajes de error en un cuadro de alerta.
        Beep
        Ret = MsgBox(ERMsg$, 1, "Haga clic en Cancelar para salir, clic en Aceptar para ignorar.")
        
        ' Si el usuaruio hace clic en Cancelar (2)...
        If Ret = 2 Then
            MSComm1.PortOpen = False    ' Cierra el puerto y sale.
        End If
        
        ' Activa el cronómetro para que el mensaje de la barra
        ' de estado se borre después de dos segundos.
        Timer2.Enabled = True
    End If
End Sub

Private Sub mnuSendText_Click()
   Dim hSend, BSize, LF&
   
   On Error Resume Next
   
   mnuSendText.Enabled = False
   tbrToolBar.Buttons("TransmitTextFile").Enabled = False
   
   ' Obtiene del usuario el nombre del archivo de texto.
   OpenLog.DialogTitle = "Enviar archivo de texto"
   OpenLog.Filter = "Archivos de texto (*.TXT)|*.txt|Todos los archivos (*.*)|*.*"
   Do
      OpenLog.CancelError = True
      OpenLog.filename = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then
        mnuSendText.Enabled = True
        tbrToolBar.Buttons("TransmitTextFile").Enabled = True
        Exit Sub
      End If
      Temp = OpenLog.filename

      ' Si el archivo no existe, vuelve.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         mnuSendText.Enabled = True
         tbrToolBar.Buttons("TransmitTextFile").Enabled = True
         Exit Sub
      End If
      If Ret Then
         Exit Do
      Else
         MsgBox Temp + " no encontrado.", 48
      End If
   Loop

   ' Abre el archivo de registro.
   hSend = FreeFile
   Open Temp For Binary Access Read As hSend
   If Err Then
      MsgBox Error$, 48
   Else
      ' Muestra el cuadro de diálogo Cancelar.
      CancelSend = False
      frmCancelSend.Label1.Caption = "Transmitiendo archivo de texto - " + Temp
      frmCancelSend.Show
      
      ' Lee el archivo en bloques del tamaño del búfer de transmisión.
      BSize = MSComm1.OutBufferSize
      LF& = LOF(hSend)
      Do Until EOF(hSend) Or CancelSend
         ' No lee demasiado al final.
         If LF& - Loc(hSend) <= BSize Then
            BSize = LF& - Loc(hSend) + 1
         End If
      
         ' Lee un bloque de datos.
         Temp = Space$(BSize)
         Get hSend, , Temp
      
         ' Transmite el bloque.
         MSComm1.Output = Temp
         If Err Then
            MsgBox Error$, 48
            Exit Do
         End If
      
         ' Espera a que se envíen todos los datos.
         Do
            Ret = DoEvents()
         Loop Until MSComm1.OutBufferCount = 0 Or CancelSend
      Loop
   End If
   
   Close hSend
   mnuSendText.Enabled = True
   tbrToolBar.Buttons("TransmitTextFile").Enabled = True
   CancelSend = True
   frmCancelSend.Hide
End Sub


' Este procedimiento establece la propiedad SThreshold, que
' determina el número máximo de caracteres que deben estar
' esperando en el búfer de salida para que se establezca
' comEvSend en la propiedad CommEvent y se dispare el evento
' OnComm.
Private Sub mnuSThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Introduzca el nuevo valor de SThreshold", "SThreshold", Str$(MSComm1.SThreshold))
    If Len(Temp) Then
        MSComm1.SThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

' Este procedimiento agrega datos a la propiedad Text del
' control Term. También filtra los caracteres de control,
' como RETROCESO, retorno de carro y avances de línea, y
' escribe datos en un archivo de registro.
' Los caracteres RETROCESO eliminan el carácter situado a
' su izquierda, ya sea en la propiedad Text o en la cadena
' pasada. Se agregan caracteres de avance de línea a todos
' los retornos de carro. El tamaño de la propiedad Text del
' control Term también se controla para que nunca exceda de
' MAXTERMSIZE caracteres.
Private Static Sub ShowData(Term As Control, Data As String)
    On Error GoTo Handler
    Const MAXTERMSIZE = 16000
    Dim TermSize As Long, i
    
    ' Se asegura que el texto existente no se haga demasiado largo.
    TermSize = Len(Term.Text)
    If TermSize > MAXTERMSIZE Then
       Term.Text = Mid$(Term.Text, 4097)
       TermSize = Len(Term.Text)
    End If

    ' Apunta al final de los datos de Term.
    Term.SelStart = TermSize

    ' Filtra y procesa los caracteres RETROCESO.
    Do
       i = InStr(Data, Chr$(8))
       If i Then
          If i = 1 Then
             Term.SelStart = TermSize - 1
             Term.SelLength = 1
             Data = Mid$(Data, i + 1)
          Else
             Data = Left$(Data, i - 2) & Mid$(Data, i + 1)
          End If
       End If
    Loop While i

    ' Elimina los avances de línea.
    Do
       i = InStr(Data, Chr$(10))
       If i Then
          Data = Left$(Data, i - 1) & Mid$(Data, i + 1)
       End If
    Loop While i

    ' Se asegura de que todos los retornos de carro tengan un
    ' avance de línea.
    i = 1
    Do
       i = InStr(i, Data, Chr$(13))
       If i Then
          Data = Left$(Data, i) & Chr$(10) & Mid$(Data, i + 1)
          i = i + 1
       End If
    Loop While i

    ' Agrega los datos filtrados a la propiedad SelText.
    Term.SelText = Data
  
    ' Registra los datos en un archivo si así se solicita.
    If hLogFile Then
       i = 2
       Do
          Err = 0
          Put hLogFile, , Data
          If Err Then
             i = MsgBox(Error$, 21)
             If i = 2 Then
                mnuCloseLog_Click
             End If
          End If
       Loop While i <> 2
    End If
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
    MsgBox Error$
    Resume Next
End Sub

Private Sub Timer2_Timer()
sbrStatus.Panels("Status").Text = "Estado:"
Timer2.Enabled = False

End Sub

' Las pulsaciones interceptadas aquí se envían
' al control MSComm, donde se devuelven a través
' del evento OnComm (comEvReceive), y se muestran
' con el procedimiento ShowData.
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    ' Si el puerto está abierto...
    If MSComm1.PortOpen Then
        ' Envía la pulsación al puerto.
        MSComm1.Output = Chr$(KeyAscii)
        
        ' Si el eco no está activado, no hay
        ' necesidad de que el control de texto
        ' muestre la tecla. Normalmente, el módem
        ' devolverá el carácter.
        If Not Echo Then
            ' Sitúa la posición al final del terminal
            txtTerm.SelStart = Len(txtTerm)
            KeyAscii = 0
        End If
    End If
     
End Sub




Private Sub tbrToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "OpenLogFile"
    Call mnuOpenLog_Click
Case "CloseLogFile"
    Call mnuCloseLog_Click
Case "DialPhoneNumber"
    Call mnuDial_Click
Case "HangUpPhone"
    Call mnuHangup_Click
Case "Properties"
    Call mnuProperties_Click
Case "TransmitTextFile"
    Call mnuSendText_Click
End Select
End Sub

Private Sub Timer1_Timer()
    ' Muestra la hora de conexión
    sbrStatus.Panels("ConnectTime").Text = Format(Now - StartTime, "hh:nn:ss") & " "
End Sub
' Llama a esta función para iniciar el cronómetro ConnectTime
Private Sub StartTiming()
    StartTime = Now
    Timer1.Enabled = True
End Sub
' Llama a esta función para detener el cronometraje
Private Sub StopTiming()
    Timer1.Enabled = False
    sbrStatus.Panels("ConnectTime").Text = ""
End Sub
