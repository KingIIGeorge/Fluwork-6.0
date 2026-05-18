VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmWave 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1380
   ClientLeft      =   1350
   ClientTop       =   4230
   ClientWidth     =   5715
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5715
   Begin VB.HScrollBar hsbWave 
      Height          =   255
      Left            =   240
      Max             =   100
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin MCI.MMControl mciWave 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
   End
   Begin VB.Label lblWaveSec 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblWave 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu AL_FILE 
      Caption         =   "&Archivo"
      Begin VB.Menu AI_OPEN 
         Caption         =   "&Abrir..."
      End
      Begin VB.Menu AI_SEPARATOR 
         Caption         =   "-"
      End
      Begin VB.Menu AI_EXIT 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const conInterval = 50
Const conIntervalPlus = 55

Dim CurrentValue As Double

Private Sub AI_EXIT_Click()
    Unload frmWave
End Sub

Private Sub AI_OPEN_Click()
    Dim msec As Double

    ' Establece el número de milisegundos entre
    ' eventos StatusUpdate sucesivos.
    mciWave.UpdateInterval = 0

    ' Muestra el cuadro de diálogo Abrir archivo.
    frmOpenDlg.dlgOpenFile.FilterIndex = 1
    frmOpenDlg.dlgOpenFile.Flags = vbOFNReadOnly Or vbOFNFileMustExist
    frmOpenDlg.dlgOpenFile.CancelError = True
    frmOpenDlg.dlgOpenFile.filename = ""
    On Error Resume Next
    frmOpenDlg.dlgOpenFile.ShowOpen

    If Err <> 0 Then
        ' Ningún archivo seleccionado en el cuadro Abrir archivo.
        Exit Sub
    End If

    ' Si el dispositivo está abierto, lo cierra.
    If Not mciWave.Mode = vbMCIModeNotOpen Then
        mciWave.Command = "Close"
    End If

    ' Abre el dispositivo con el nuevo nombre de archivo.
    mciWave.filename = frmOpenDlg.dlgOpenFile.filename
    On Error GoTo MCI_ERROR
    mciWave.Command = "Open"
    On Error GoTo 0
    Caption = DialogCaption + mciWave.filename
    
    ' Establece las etiquetas de temporización en el formuario.
    mciWave.TimeFormat = vbMCIFormatMilliseconds
    lblWave.Caption = "0.0"
    msec = (CDbl(mciWave.Length) / 1000)
    lblWaveSec.Caption = Format$(msec, "0.00")

    ' Establece los valores de la barra de desplazamiento.
    hsbWave.value = 0
    CurrentValue = 0#
    Exit Sub

MCI_ERROR:
    DisplayErrorMessageBox
    Resume MCI_EXIT

MCI_EXIT:
    Unload frmWave
    
End Sub

Private Sub Form_Load()
    ' Obliga al control MCI multimedia a terminar antes de volver
    ' a la aplicación.
    frmWave.mciWave.Wait = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMCITest.Show
End Sub

Private Sub mciWave_PauseClick(Cancel As Integer)
    ' Establece el número de milisegundos entre
    ' eventos StatusUpdate sucesivos.
    mciWave.UpdateInterval = 0
End Sub

Private Sub mciWave_PlayClick(Cancel As Integer)
    ' Establece el número de milisegundos entre
    ' eventos StatusUpdate sucesivos.
    mciWave.UpdateInterval = conInterval
End Sub

Private Sub mciWave_PrevClick(Cancel As Integer)
    ' Establece el número de milisegundos entre
    ' eventos StatusUpdate sucesivos.
    mciWave.UpdateInterval = 0
    
    ' Restablece los valores de barra de desplazamiento.
    hsbWave.value = 0
    CurrentValue = 0#
    
    mciWave.Command = "Prev"
End Sub

Private Sub mciWave_StatusUpdate()
    Dim value As Integer

    ' Si el dispositivo no está reproduciendo,
    ' restablece el principio.
    If Not mciWave.Mode = vbMCIModePlay Then
        hsbWave.value = hsbWave.Max
        mciWave.UpdateInterval = 0
        Exit Sub
    End If
    
    ' Determina qué parte del archivo se ha reproducido.
    ' Establece un valor entre 0 y 100 en la barra de 
    ' desplazamiento.
    CurrentValue = CurrentValue + conIntervalPlus
    value = CInt((CurrentValue / mciWave.Length) * 100)
    
    If value > hsbWave.Max Then
        value = 100
    End If

    hsbWave.value = value
End Sub

