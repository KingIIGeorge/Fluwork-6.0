VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cajero automático"
   ClientHeight    =   3765
   ClientLeft      =   1455
   ClientTop       =   2460
   ClientWidth     =   5070
   Icon            =   "openbank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5070
   Begin VB.Frame Label2 
      Caption         =   "Seleccione un idioma"
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "English"
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Français"
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Deutsch"
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Italiano"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Español"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Japan"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image 
         Height          =   375
         Index           =   0
         Left            =   360
         Picture         =   "openbank.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image Image 
         Height          =   375
         Index           =   1
         Left            =   360
         Picture         =   "openbank.frx":0316
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image 
         Height          =   375
         Index           =   2
         Left            =   360
         Picture         =   "openbank.frx":0620
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Image 
         Height          =   375
         Index           =   3
         Left            =   2520
         Picture         =   "openbank.frx":092A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image Image 
         Height          =   375
         Index           =   4
         Left            =   2520
         Picture         =   "openbank.frx":0C34
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image 
         Height          =   375
         Index           =   5
         Left            =   2520
         Picture         =   "openbank.frx":0F3E
         Stretch         =   -1  'True
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   """Servicio personalizado en todo el mundo"""
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Top             =   840
      Width           =   4700
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Cajero automático"
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   600
      Width           =   4700
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Bienvenido al"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   4700
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "BANCO MUNDIAL VIAJERO"
      Height          =   240
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   360
      Width           =   4700
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_Base = "0{46B50776-F609-11CF-8A54-0020AF939EAB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cada botón de comando representa un idioma. Si el usuario elige inglés,
'la aplicación cargará recursos con números de identificación del 16 al 47.
Option Explicit

Sub command1_click()
    i = 16              ' El usuario eligió inglés.
    frmInput.Show vbModal
End Sub

Sub Command2_Click()
    i = 48              ' El usuario eligió francés.
    frmInput.Show vbModal
End Sub

Sub Command3_Click()
    i = 80              ' El usuario eligió alemán.
    frmInput.Show vbModal
End Sub

Sub Command4_Click()
    i = 112              ' El usuario eligió italiano.
    frmInput.Show vbModal
End Sub

Sub Command5_Click()
    i = 144              ' El usuario eligió español.
    frmInput.Show vbModal
End Sub

Sub Command6_Click()     ' El japonés sólo esta disponible en sistemas DBCS
    i = 176              ' El usuario eligió japonés.
    frmInput.Show vbModal
End Sub

Private Sub Form_Load()
    Dim iIndex As Byte

    ' Inicializa la tabla de conversión actual.
    ConversionTable_Initialize
    
    ' Inicializa el cursor del icono del ratón sobre el botón.
    Cursor_Initialize
    SetCursor Command1
    SetCursor Command2
    SetCursor Command3
    SetCursor Command4
    SetCursor Command5
    SetCursor Command6

    ' Para DBCS, VB debe establecer el nombre de la fuente al cargar
    ' el proyecto. Para que este ejemplo funcione en cualquier sistema
    ' las propiedades de la fuente se establecen por programa.
    For iIndex = 0 To 3
        Label(iIndex).FontSize = 10
        Label(iIndex).FontBold = True
    Next iIndex
    
    ' Con DBCS, los caracteres ANSI superiores no se pueden leer con sistemas
    ' orientales. Como algunos de estos caracteres se usan en los botones de 
    ' idiomas, es necesario cambiarlos en este caso.
    If IMEStatus = vbIMENoOp Then
        Command2.Caption = "Français"
        Command5.Caption = "Español"
    End If
End Sub

