VERSION 5.00
Begin VB.Form statsel 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "statsel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000C0C0&
      Caption         =   "DIAGNOSTIC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "DEPOSITO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00004000&
      Caption         =   "ANULADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRESUP."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF0000&
      Caption         =   "ENTREGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LISTA BRGS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000C000&
      Caption         =   "LISTA NR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF00&
      Caption         =   "PV. EXT."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000C0&
      Caption         =   "REP. EXT."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Caption         =   "CHEQUEO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ENTREGADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF00FF&
      Caption         =   "STD/BY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&ACEPTAR"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FF00&
      Caption         =   "LISTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000FF&
      Caption         =   "REPARANDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "POR VER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   1800
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   1920
      X2              =   4200
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      X1              =   4200
      X2              =   4200
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   1920
      X2              =   4200
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Esto le permitira poder realizar Busquedas Avanzadas."
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2040
      TabIndex        =   18
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Seleccione el estado   del equipo en cuestion."
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Para cancelar la seleccion de estado existente solo debe cerrar esta ventana."
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2040
      TabIndex        =   16
      Top             =   960
      Width           =   2055
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   1800
      X2              =   1800
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   1800
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "statsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Byte
Dim statusColor As Long

For i = 0 To StatusCatalog.STATUS_COUNT - 1
If Option1(i).Value = True Then
StatusCatalog.ApplyStatusSelection i + 1
statusColor = StatusCatalog.ColorForStatus(registro.estado)
If statusColor <> -1 Then Form1.lbllista.ForeColor = statusColor
End If
Next i
Unload Me
End Sub
