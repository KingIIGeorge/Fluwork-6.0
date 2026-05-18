VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de BIBLIO"
   ClientHeight    =   1965
   ClientLeft      =   1995
   ClientTop       =   4455
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1965
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2820
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Copyright (c) 1995 Microsoft Corporation"
      Height          =   255
      Left            =   1980
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Aplicación de ejemplo del control Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "BIBLIO -- Examinador de base de datos bibliográfica"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6555
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload frmAbout
End Sub


