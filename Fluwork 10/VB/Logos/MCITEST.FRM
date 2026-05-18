VERSION 5.00
Begin VB.Form frmMCITest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicación del control MCI"
   ClientHeight    =   870
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   3765
   Icon            =   "mcitest.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3765
   Begin VB.Menu AL_DEVICES 
      Caption         =   "&Dispositivos"
      Begin VB.Menu AI_CDAUDIO 
         Caption         =   "&CD Audio..."
      End
      Begin VB.Menu AI_SEQUENCER 
         Caption         =   "&Secuenciador MIDI..."
      End
      Begin VB.Menu AI_WAVEAUDIO 
         Caption         =   "&Archivo de sonido"
      End
   End
   Begin VB.Menu AL_HELP 
      Caption         =   "Ay&uda"
      Begin VB.Menu AI_ABOUT 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMCITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AI_ABOUT_Click()
    ' Muestra el cuadro Acerca de. Permite procesar
    ' los mensajes de Windows antes de establecer el
    ' enfoque en el botón Aceptar.
    frmAboutBox.Show
    WaitForEventsToFinish 50
    On Error Resume Next
    frmAboutBox.BTN_OK.SetFocus
End Sub

Private Sub AI_CDAUDIO_Click()
    frmCD.Show
    frmMCITest.Hide
    DisplayWarning
    
End Sub

Private Sub AI_SEQUENCER_Click()
    DialogCaption = "Secuenciador MDI - "
    frmWave.Caption = "Secuenciador MIDI"
    
    ' Establece el cuadro de diálogo Abrir archivo para que
    ' muestre los archivos *.mid.
    frmOpenDlg.dlgOpenFile.Filter = "Archivo MIDI (*.mid)|*.mid"
    frmWave.mciWave.DeviceType = "Sequencer"
    frmWave.Show
    frmMCITest.Hide
End Sub

Private Sub AI_WAVEAUDIO_Click()
    DialogCaption = "Archivo de sonido - "
    frmWave.Caption = "Archivo de sonido"
    
    ' Establece el cuadro de diálogo Abrir archivo para que
    ' muestre los archivos *.wav.
    frmOpenDlg.dlgOpenFile.Filter = "Archivo de audio (*.wav)|*.wav"
    frmWave.mciWave.DeviceType = "WaveAudio"
    frmWave.Show
    frmMCITest.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Recuerde descargar el formulario OpenDlg, ya que
    ' sólo se "oculta" al cerrarlo.
    Unload frmOpenDlg
End Sub

