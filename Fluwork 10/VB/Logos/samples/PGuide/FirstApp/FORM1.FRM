VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Librería"
   ClientHeight    =   5490
   ClientLeft      =   330
   ClientTop       =   1305
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "..\..\..\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "All Titles"
      Top             =   4440
      Width           =   6375
   End
   Begin VB.ListBox List1 
      Height          =   1260
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   6375
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0010
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    List1.Clear         ' Borra el cuadro de lista
End Sub

Private Sub Command2_Click()
    Unload Form1
    End                 ' Termina la aplicación
End Sub

Private Sub Data1_Reposition()
    ' Asigna la propiedad Title al título del control
    Data1.Caption = Data1.Recordset("Title")
End Sub

Private Sub DBGrid1_DblClick()
    ' Agrega el título a la lista
    List1.AddItem Data1.Recordset("Title")
End Sub
