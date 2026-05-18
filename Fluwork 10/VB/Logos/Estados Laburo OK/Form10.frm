VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione destino de archivo"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "email.csv"
      Top             =   360
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H8000000B&
      Height          =   1650
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H8000000A&
      Height          =   1665
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unida&des:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "c:\"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Carpetas:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Nombre del archivo:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
nombrecompleto = AddBackslash(directorio) & Trim(Text1.Text)
Unload Me
Sleep 1000
End Sub

Private Sub Command2_Click()
Unload Me
nombrecompleto = ""
End Sub

Private Sub Dir1_Change()
directorio = Dir1.Path
ChDir Trim(directorio)
File1.Path = directorio
If Len(directorio) < 20 Then
Label3.Caption = directorio
Else
Label3.Caption = ".."
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo drivenotready
unidad = Drive1.Drive
ChDrive (unidad)
Dir1.Path = unidad

drivenotready:
If Err.Number = 68 Then
If MsgBox((("No se puede tener acceso a ") & StrConv(unidad, 1) & ":\"), vbOKOnly, "Fluwork") = vbOK Then
End If
End If
End Sub


Private Sub Form_Load()
unidad = Drive1.Drive
directorio = Dir1.Path

End Sub

