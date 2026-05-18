VERSION 5.00
Begin VB.Form AlarmForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reloj con alarma"
   ClientHeight    =   780
   ClientLeft      =   5010
   ClientTop       =   4245
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ALARM.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   780
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2640
      Top             =   120
   End
   Begin VB.Label lblTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "AlarmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AlarmTime
Const conMinimized = 1


Private Sub Form_Click()
    AlarmTime = InputBox("Escriba la hora de la alarma", "Alarma de VB", AlarmTime)
    If AlarmTime = "" Then Exit Sub
    If Not IsDate(AlarmTime) Then
        MsgBox "La hora no es válida."
    Else                                    ' La cadena devuelta por InputBox es una hora válida,
        AlarmTime = CDate(AlarmTime)        ' de modo que se almacena en AlarmTime como Fecha/hora.
    End If
End Sub

Private Sub Form_Load()
    AlarmTime = ""
End Sub

Private Sub Form_Resize()
    If WindowState = conMinimized Then      ' Si el formulario se minimiza, presenta la hora en un título.
        SetCaptionTime
    Else
        Caption = "Reloj con alarma"
    End If
End Sub

Private Sub SetCaptionTime()
    Caption = Format(Time, "Medium Time")   ' Presenta la hora con el formato Medium Time.
End Sub

Private Sub Timer1_Timer()
Static AlarmSounded As Integer
    If lblTime.Caption <> CStr(Time) Then
        ' Ahora es un segundo diferente del presentado.
        If Time >= AlarmTime And Not AlarmSounded Then
            Beep
            MsgBox "Alarma a las " & Time, , "Alarma"
            AlarmSounded = True
        ElseIf Time < AlarmTime Then
            AlarmSounded = False
        End If
        If WindowState = conMinimized Then
            ' Si está minimizado, actualiza el título del formulario cada minuto.
            If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
        Else
            ' Si no, actualiza la etiqueta del formulario cada segundo.
            lblTime.Caption = Time
        End If
    End If
End Sub

