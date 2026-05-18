VERSION 5.00
Begin VB.Form frmVCR 
   BackColor       =   &H00000000&
   Caption         =   "VBTV"
   ClientHeight    =   7125
   ClientLeft      =   420
   ClientTop       =   525
   ClientWidth     =   7110
   FillStyle       =   0  'Solid
   Icon            =   "vcr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7110
   Begin VB.CommandButton cmdSet 
      Caption         =   "Establecer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4500
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Establezca la hora de grabación"
      Top             =   6000
      Width           =   1155
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Left            =   6240
      Top             =   3240
   End
   Begin VB.Timer tmr1 
      Interval        =   65535
      Left            =   6240
      Top             =   2640
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Abajo"
      Height          =   325
      Left            =   6120
      MaskColor       =   &H00000000&
      TabIndex        =   11
      ToolTipText     =   "Selector de canal"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Arriba"
      Height          =   325
      Left            =   6120
      MaskColor       =   &H00000000&
      TabIndex        =   10
      ToolTipText     =   "Selector de canal"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Expulsar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      MaskColor       =   &H00000000&
      TabIndex        =   7
      ToolTipText     =   "Salir de VBTV"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "ll"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1005
      MaskColor       =   &H00000000&
      TabIndex        =   6
      ToolTipText     =   "Pausa"
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdRec 
      Caption         =   "Grab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      MaskColor       =   &H00000000&
      TabIndex        =   5
      ToolTipText     =   "Grabar"
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2955
      MaskColor       =   &H00000000&
      TabIndex        =   4
      ToolTipText     =   "Avance rápido"
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdRewind 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2310
      MaskColor       =   &H00000000&
      TabIndex        =   3
      ToolTipText     =   "Rebobinar"
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Parar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   2
      ToolTipText     =   "Parar"
      Top             =   6000
      Width           =   615
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   360
      ScaleHeight     =   4215
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      MaskColor       =   &H00000000&
      TabIndex        =   0
      ToolTipText     =   "Reproducir"
      Top             =   6000
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   7080
      Y1              =   4660
      Y2              =   4660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7080
      Y1              =   4700
      Y2              =   4700
   End
   Begin VB.Shape shpPlay 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   615
      Shape           =   3  'Circle
      Top             =   5835
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpForward 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3210
      Shape           =   3  'Circle
      Top             =   5835
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpRewind 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   2565
      Shape           =   3  'Circle
      Top             =   5835
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpRec 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   1905
      Shape           =   3  'Circle
      Top             =   5835
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpPause 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   1260
      Shape           =   3  'Circle
      Top             =   5835
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image img2 
      Height          =   1155
      Left            =   1680
      Picture         =   "vcr.frx":0442
      Top             =   6720
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image img1 
      Height          =   1155
      Left            =   240
      Picture         =   "vcr.frx":10CC
      Top             =   6720
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lblBrand 
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Visual Basic VCR - Versión 1.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Image imgTapeSlot 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2640
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Mostrar canal"
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmVCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
' Propósito:  Formulario principal de la aplicación de
' ejemplo VCR. Emula una grabadora de vídeo
' mediante objetos de Visual Basic.
'*******************************************************
Option Explicit

' Crea una instancia de la clase Tape
Dim Tape As New clsTape

Dim vntChannel As Variant   ' Número de canal

' Constantes para la función QBColor
Const vcrBlack = 0
Const vcrGreen = 2
Const vcrCyan = 3
Const vcrRed = 4
Const vcrMagenta = 5
Const vcrYellow = 6
Const vcrWhite = 7
Const vcrGray = 8
Const vcrLightBlue = 9
Const vcrLightGreen = 10
Const vcrLightCyan = 11
Const vcrLightRed = 12
Const vcrLightMagenta = 13

Private Sub cmdDown_Click()
    ' si está dentro del intervalo, establece el número de canal
    If vntChannel > 2 Then
        vntChannel = vntChannel - 1
    Else
        vntChannel = 13
    End If
    
    ' asigna la variable del canal a la presentación en pantalla
    lblChannel.Caption = vntChannel
End Sub

Private Sub cmdExit_Click()
    ' descarga el formulario y libera la referencia
    Unload Me
    Set frmVCR = Nothing
End Sub

Private Sub cmdForward_Click()
    ' llama a la función que guarda el canal anterior
    SaveChannel vntChannel
    ' para reproducir la cinta tiene que estar en el canal 3
    vntChannel = 3
    lblChannel.Caption = vntChannel
    ' Establece las propiedades de la clase Tape
    Tape.Forward = True
    Tape.Speed = 50
    ' Inicia el temporizador
    tmr2.Enabled = True
    tmr2.Interval = Tape.Speed
    ' Llama a la función que actualiza los controles
    ButtonManager frmVCR.cmdForward
End Sub

Private Sub cmdPause_Click()
    ' Detiene el temporizador
    tmr2.Enabled = False
    ' Llama a la función que actualiza los controles
    ButtonManager frmVCR.cmdPause
End Sub

Private Sub cmdPlay_Click()
    ' llama a la función que guarda el canal anterior
    SaveChannel vntChannel
    ' para reproducir la cinta tiene que estar en el canal 3
    vntChannel = 3
    lblChannel.Caption = vntChannel
    ' Establece las propiedades de la clase Tape
    Tape.Forward = True
    Tape.Speed = 300
    ' Inicia el temporizador
    tmr2.Enabled = True
    tmr2.Interval = Tape.Speed
    ' Llama a la función que actualiza los controles
    ButtonManager frmVCR.cmdPlay
End Sub

Private Sub cmdRec_Click()
    Dim strStatus As String     'Presenta texto
    
    ' llama a la función que guarda el canal anterior
    SaveChannel vntChannel
    ' Borra la presentación
    picTV.Cls
    ' Presenta el estado
    strStatus = "Grabando: Canal " & vntChannel
    picTV.Print strStatus
    strStatus = lblTime.Caption
    picTV.Print strStatus
    ' Llama a la función que actualiza los controles
    ButtonManager frmVCR.cmdRec
End Sub

Private Sub cmdRewind_Click()
    ' llama a la función que guarda el canal anterior
    SaveChannel vntChannel
    ' para reproducir la cinta tiene que estar en el canal 3
    vntChannel = 3
    lblChannel.Caption = vntChannel
    ' Establece las propiedades de la clase Tape
    Tape.Forward = False
    Tape.Speed = 50
    ' Inicia el temporizador
    tmr2.Enabled = True
    tmr2.Interval = Tape.Speed
    ' Llama a la función que actualiza los controles
    ButtonManager frmVCR.cmdRewind
End Sub

Private Sub cmdSet_Click()
    ' muestra el formulario de entrada del usuario en modo modal
    frmSetTime.Show vbModal
End Sub

Private Sub cmdStop_Click()
    Dim intChannel As Integer   'Número de canal
    
    ' Detiene el temporizador
    tmr2.Enabled = False
    ' Llama a la función que actualiza los controles
    ButtonManager frmVCR.cmdStop
    ' Borra la presentación
    picTV.Cls
    ' restablece el canal anterior guardado
    intChannel = SaveChannel(0)
    vntChannel = intChannel
    lblChannel.Caption = vntChannel
End Sub

Private Sub cmdUp_Click()
    ' si está dentro del intervalo, establece el número de canal
    If vntChannel < 13 Then
        vntChannel = vntChannel + 1
    Else
        vntChannel = 2
    End If
    
     ' asigna la variable de canal a la presentación de la pantalla
     lblChannel.Caption = vntChannel
End Sub

Private Sub Form_Load()
    ' Muestra la hora actual
    lblTime.Caption = Format((Now), "h:mm AM/PM")
    ' Establece la altura del formulario
    frmVCR.Height = 6990
    img1.Visible = True
    ' Presenta el formulario
    Me.Show
    ' establece el canal inicial
    vntChannel = 3
    lblChannel.Caption = vntChannel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Libera las referencias
    Set Tape = Nothing
    Set Recorder = Nothing
    Set frmVCR = Nothing
End Sub

Private Sub lblChannel_Change()
    ' Cambia el color de presentación de acuerdo con el canal
    Select Case vntChannel
        Case 2
            picTV.BackColor = QBColor(vcrGreen)
        Case 3
            picTV.BackColor = QBColor(vcrWhite)
        Case 4
            picTV.BackColor = QBColor(vcrRed)
        Case 5
            picTV.BackColor = QBColor(vcrMagenta)
        Case 6
            picTV.BackColor = QBColor(vcrYellow)
        Case 7
            picTV.BackColor = QBColor(vcrCyan)
        Case 8
            picTV.BackColor = QBColor(vcrGray)
        Case 9
            picTV.BackColor = QBColor(vcrLightBlue)
        Case 10
            picTV.BackColor = QBColor(vcrLightGreen)
        Case 11
            picTV.BackColor = QBColor(vcrLightCyan)
        Case 12
            picTV.BackColor = QBColor(vcrLightRed)
        Case 13
            picTV.BackColor = QBColor(vcrLightMagenta)
    End Select
    
    ' Borra la presentación
    picTV.Cls
    ' Presenta el canal y la hora
    picTV.Print "Canal: " & vntChannel
    picTV.Print lblTime.Caption
End Sub

Private Sub tmr1_Timer()
    ' Actualiza la presentación de la hora
    lblTime.Caption = Format((Now), "h:mm AM/PM")
    ' Si la propiedad de Recorder está activada
    If Recorder.Enabled = True Then
        ' Es hora de grabar
        If Recorder.StartRecording = lblTime.Caption Then
            ' Empieza la grabación
            vntChannel = Recorder.Channel
            lblChannel.Caption = vntChannel
            ' Activa el botón Record
            cmdRec.Value = True
            ' borra la propiedad de la clase Recorder
            Recorder.StartRecording = Empty
        End If
    Else
        ' Si es la hora de terminar la grabación
        If Recorder.StopRecording = lblTime.Caption Then
            ' Activa el botón Stop
            cmdStop.Value = True
            ' borra la propiedad de la clase Recorder
            Recorder.StopRecording = Empty
        End If
    End If
End Sub


Private Sub tmr2_Timer()
    Dim intWidth As Integer     'Anchura
    Dim intLeft As Integer      'Posición izquierda
    Dim objImage As Control     'Control Image
    
    ' Obtiene la anchura de la presentación
    intWidth = picTV.Width
    ' Llama al método de la clase Tape
    ' para reproducir la cinta.
    Tape.Animate intWidth
    
    ' Obtiene la propiedad Left a partir de la clase
    intLeft = Tape.Left
    ' Muestra la primera o la segunda imagen
    If img1.Visible = True Then
        img1.Visible = False
        Set objImage = img2
    Else
        img1.Visible = True
        Set objImage = img1
    End If
    
    ' Borra la presentación
    picTV.Cls
    ' Muestra la nueva imagen en la nueva posición
    picTV.PaintPicture objImage.Picture, intLeft, 1200
End Sub



