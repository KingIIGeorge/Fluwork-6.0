VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmBookSales 
   Caption         =   "Servidor de venta de libros"
   ClientHeight    =   2235
   ClientLeft      =   840
   ClientTop       =   7170
   ClientWidth     =   5520
   ClipControls    =   0   'False
   Icon            =   "Book_svr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   5520
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FontSize        =   1.17485e-38
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contador de instancias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   180
      Width           =   2025
   End
   Begin VB.Label lblInstanceCount 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   750
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   765
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   1380
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta local:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   765
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   510
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Petición remota:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   0
      Top             =   525
      Width           =   1425
   End
End
Attribute VB_Name = "frmBookSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

