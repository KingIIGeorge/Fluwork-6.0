VERSION 5.00
Begin VB.Form frmGeoFacts 
   BackColor       =   &H00C0C0C0&
   Caption         =   "GeoFacts"
   ClientHeight    =   4440
   ClientLeft      =   1890
   ClientTop       =   1440
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   5805
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox lstTopRanking 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2430
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cmbFeatures 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox cmbContinents 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2412
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Características del continente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label IvoSa 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Continente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label IvoWilliamSalmre 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GeoFacts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2772
   End
End
Attribute VB_Name = "frmGeoFacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Load()
    Setup   ' Establece las variables objeto a los objetos de Excel.
    FillContinentsList ' Llena los continentes.
End Sub

Sub Form_Unload(Cancel As Integer)
    CleanUp ' Establece las variables objeto a Nothing.
End Sub

Private Sub cmbContinents_Change()
    FillFeaturesList    ' Llena el cuadro combinado Caracteristicas del continente.
End Sub

Private Sub cmbFeatures_Change()
    FillTopRankingList ' Llena el cuadro de lista lstTopRanking.
End Sub

Sub cmbFeatures_Click()
    FillTopRankingList ' Llena el cuadro de lista lstTopRanking.
End Sub

Sub cmbContinents_Click()
    FillFeaturesList    ' Llena el cuadro combinado Caracteristicas del continente.
End Sub

