VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Sugerencia del día"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5850
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Mostrar sugerencias al iniciar"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2475
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Siguiente sugrencia"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "¿Sabía que..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' La base de datos de sugerencias.
Dim Tips As New Collection

' Nombre del archivo de sugerencias
Const TIP_FILE = "TIPOFDAY.TXT"

' Índice de la colección con la sugerencia visualizada actualmente.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Selecciona una sugerencia aleatoriamente.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' O, recorre las sugerencias por orden

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Muestra la sugerencia.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Se lee cada sugerencia desde el archivo.
    Dim InFile As Integer   ' Descriptor del archivo.
    
    ' Obtiene el siguiente descriptor libre.
    InFile = FreeFile
    
    ' Comprueba que se ha especificado un archivo.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Comprueba que existe un archivo antes de abrirlo.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Lee la colección desde el archivo de texto.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Muestra una sugerencia aleatoriamente.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' guarda si se debe o no mostrar el formulario al iniciar
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    
    ' Comprueba si se debe mostrar al iniciar
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
        
    ' Establece la casilla de verificación, esto obliga a que el valor se vuelva a escribir en el Registro
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Semilla aleatoria
    Randomize
    
    ' Lee el archivo de sugerencias y muestra una de forma aleatoria
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "No se encontró el archivo " & TIP_FILE & "." & vbCrLf & vbCrLf & _
           "Cree un archivo de texto llamado " & TIP_FILE & " usando el Bloc de notas con una sugerencia por línea. " & _
           "Sitúelo en el mismo directorio que la aplicación. "
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
