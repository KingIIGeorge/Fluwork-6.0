VERSION 5.00
Begin VB.Form frmSetTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   780
   ClientTop       =   1590
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtChannel 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Text            =   "3"
      ToolTipText     =   "Escriba el canal a grabar"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtEnd 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Text            =   "12:30 AM"
      ToolTipText     =   "Escriba la hora a la que se para de grabar"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtStart 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Text            =   "12:00 AM"
      ToolTipText     =   "Escriba la hora a la que se comienza a grabar"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Canal"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1980
      Width           =   795
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Hora final"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Hora inicial"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "frmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' Propósito:  Formulario para obtener la entrada del usuario que establece
'           la hora y el canal de grabación.
'           Forma parte de la aplicación de ejemplo VCR
'***************************************************************************
Option Explicit

Private Sub cmdCancel_Click()
    ' descarga el formulario, libera su referencia
    Unload Me
    Set frmSetTime = Nothing
End Sub

Private Sub cmdOK_Click()
    Dim intOK As Integer        'Valor devuelto
    Dim intChannel As Integer   'Número de canal
    Dim strMsg As String        'Texto del mensaje
    
    ' Llama a la función que valida la entrada
    intOK = ValidateTime(txtStart.Text, txtStart)
    ' Si no es válida, sale
    If intOK = False Then Exit Sub
        
    ' Llama a la función que valida la entrada
    intOK = ValidateTime(txtEnd.Text, txtEnd)
    ' Si no es válida, sale
    If intOK = False Then Exit Sub
    
    ' Lee el valor del canal presentado en la pantalla
    intChannel = Val(txtChannel)
    ' Si está fuera de intervalo, advierte al usuario
    If intChannel < 2 Or intChannel > 13 Then
        strMsg = "El canal debe estar entre 2 y 13"
        MsgBox strMsg, vbOKOnly, "Canal no válido"
        txtChannel.SetFocus
        Exit Sub
    End If
    
    ' Establece las propiedades de la clase Recorder
    Recorder.StartRecording = CVar(txtStart.Text)
    Recorder.StopRecording = CVar(txtEnd.Text)
    Recorder.Channel = intChannel
        
    ' descarga el formulario, libera la referencia
    Unload Me
    Set frmSetTime = Nothing
End Sub

Private Sub Form_Load()
    ' Si la propiedad de la clase contiene un
    ' valor, lo asigna al cuadro de texto
    If Recorder.StartRecording <> Empty Then
        txtStart.Text = Recorder.StartRecording
    End If
    If Recorder.StopRecording <> Empty Then
        txtEnd.Text = Recorder.StopRecording
    End If
    txtChannel = Recorder.Channel
End Sub
'**********************************************************
' Propósito:  Valida la cadena de hora introducida por el
'           usuario. Si es correcta, da formato a la
'           cadena, si no, presenta un mensaje.
' Entradas:   Time:   La cadena pasada desde un
'                   cuadro de texto.
'           Field:  El nombre del control cuadro de texto.
' Devuelve:  True si la hora es válida, False si la
'           cadena no es una hora válida.
'**********************************************************
Function ValidateTime(Time As String, Field As Control) As Boolean
    Dim strMsg As String    ' Texto del mensaje
    
    ' La función IsDate verifica si la hora es válida
    If IsDate(Time) Then
        ' Formatea la cadena y devuelve True
        Field.Text = Format$(Time, "h:mm AM/PM")
        ValidateTime = True
    Else
        ' Notifica el error al usuario
        strMsg = "Introduzca una hora válida "
        strMsg = strMsg & "(Hora:Minuto AM/PM)"
        MsgBox strMsg, vbOKOnly, "Hora no válida"
        ' Devuelve el enfoque al cuadro de texto
        Field.SetFocus
        ' Devuelve False
        ValidateTime = False
    End If
    
End Function

Private Sub txtChannel_GotFocus()
    ' Selecciona el texto existente
    txtChannel.SelStart = 0
    txtChannel.SelLength = Len(txtChannel.Text)
End Sub

Private Sub txtChannel_KeyPress(KeyAscii As Integer)
    ' Utiliza la función IsNumeric para asegurar que
    ' sólo se introducen números
    If Not IsNumeric(Chr(KeyAscii)) Then
        ' No es un número, así que lo establece a nulo
        KeyAscii = 0
    End If
End Sub

Private Sub txtEnd_GotFocus()
    ' Selecciona el texto existente
    txtEnd.SelStart = 0
    txtEnd.SelLength = Len(txtEnd.Text)
End Sub

Private Sub txtEnd_LostFocus()
    Dim intOK As Integer    ' Valor devuelto
    
    ' Llama a la función que valida la entrada
    intOK = ValidateTime(txtEnd.Text, txtEnd)
End Sub

Private Sub txtStart_GotFocus()
    ' Selecciona el texto existente
    txtStart.SelStart = 0
    txtStart.SelLength = Len(txtStart.Text)
End Sub

Private Sub txtStart_LostFocus()
    Dim intOK As Integer    'Valor devuelto
    
    ' Llama a la función que valida la entrada
    intOK = ValidateTime(txtStart.Text, txtStart)
End Sub
