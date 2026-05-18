VERSION 5.00
Begin VB.Form frmErrors 
   Caption         =   "Ejemplo de tratamiento de errores"
   ClientHeight    =   2670
   ClientLeft      =   1140
   ClientTop       =   1530
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeError 
      Caption         =   "Generar un error"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton cmdCentral 
      Caption         =   "Tratamiento de errores centralizado"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtFileSpec 
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Text            =   "*:\error.xyz"
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdInline 
      Caption         =   "Tratamiento de errores inmediato"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Sin tratamiento de errores"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Escriba un nombre de archivo no válido:"
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCentral_Click()
    Dim strErrnum As String
    Dim intErr As Integer
    Dim intReturn As Integer
    
    On Error GoTo CallFileErrors
    strErrnum = InputBox("Escriba un número de error", "Errores", "68")
    intErr = Val(strErrnum)
    Err.Raise Number:=intErr
    Exit Sub
CallFileErrors:
    intReturn = FileErrors()
    If intReturn = 0 Then
        Resume
    ElseIf intReturn = 1 Then
        Resume Next
    ElseIf intReturn = 2 Then
        MsgBox "Error no recuperable"
        End
    Else
        MsgBox "Error desconocido"
        Resume Next
    End If
        
End Sub

Private Sub cmdInline_Click()
    Dim blnResult As Boolean
    
    blnResult = FileExists2(txtFileSpec.Text)
End Sub


Private Sub cmdMakeError_Click()
    Err.Raise Number:=71        ' Simula el error "El disco no está listo".
End Sub

Private Sub cmdNone_Click()
    Dim blnResult As Boolean
    
    blnResult = FileExists1(txtFileSpec.Text)
End Sub

Function FileExists1(filename) As Boolean
    ' Sin control de errores
    FileExists1 = (Dir(filename) <> "")
End Function
Function FileExists2(filename) As Boolean
    Dim Msg As String
    ' Activa la interceptación de errores para que el controlador de errores responda
    ' cuando se produzca un error.
    On Error GoTo CheckError
        FileExists2 = (Dir(filename) <> "")
        ' Evita la ejecución del controlador de errores si no los hay.
        Exit Function

CheckError:                 ' Salta aquí si se produce un error.
    ' Define constantes que representan códigos de error internos de
    ' Visual Basic.
    Const mnErrDiskNotReady = 71, mnErrDeviceUnavailable = 68
    ' vbExclamation, vbOK, vbCancel, vbCritical y vbOKCancel son
    'constantes definidas en la biblioteca de tipos de VBA.
    If (Err.Number = mnErrDiskNotReady) Then
        Msg = "Introduzca correctamente el disquete en la unidad."
        ' Presenta un cuadro de mensaje con un icono de exclamación y
        ' los botones Aceptar y Cancelar.
        If MsgBox(Msg, vbExclamation & vbOKCancel) = vbOK Then
            Resume
        Else
            Resume Next
        End If
    ElseIf Err.Number = mnErrDeviceUnavailable Then
        Msg = "Esta unidad o ruta no existe: " & filename
        MsgBox Msg, vbExclamation
        Resume Next
    Else
        Msg = "Error no esperado Nº" & Str(Err.Number) & " ocurrió: " _
        & Err.Description
        ' Presenta un cuadro de mensaje con un icono de parada y un botón Aceptar.
        MsgBox Msg, vbCritical
        Stop
    End If
    Resume
End Function

