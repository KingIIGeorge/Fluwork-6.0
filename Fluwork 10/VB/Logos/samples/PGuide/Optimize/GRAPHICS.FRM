VERSION 5.00
Begin VB.Form Graphics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gráficos"
   ClientHeight    =   6585
   ClientLeft      =   600
   ClientTop       =   690
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "GRAPHICS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   439
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   Tag             =   "Resource"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   7560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   7620
      Begin VB.CommandButton cmdRedraw 
         Caption         =   "&Volver a dibujar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   5400
         TabIndex        =   2
         Top             =   480
         Width           =   1812
      End
      Begin VB.ListBox lstRedrawType 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3420
         TabIndex        =   1
         Top             =   30
         Width           =   1755
         Begin VB.OptionButton optGraphicType 
            Caption         =   "Metarchivo"
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
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   555
            Width           =   1455
         End
         Begin VB.OptionButton optGraphicType 
            Caption         =   "Mapa de bits"
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
            Left            =   120
            TabIndex        =   3
            Top             =   255
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Label lblTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image imgMetafile 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   135
         Picture         =   "GRAPHICS.frx":030A
         Top             =   405
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Image imgBitmap 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   1170
         Picture         =   "GRAPHICS.frx":0540
         Top             =   285
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox Container 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   367
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   506
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1335
      Width           =   7620
      Begin VB.Image cell 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Index           =   0
         Left            =   0
         Picture         =   "GRAPHICS.frx":18AA
         Top             =   240
         Width           =   1110
      End
   End
End
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Investigue los distintos procedimientos ReSizeCells generales.
' Intente también modificar la propiedad Stretch del control Cell.

Dim iFormLoad As Integer

Private Sub cmdRedraw_Click()
    ReDraw

End Sub

Private Sub Form_Load()
  Dim i As Integer
    
  ' Indica que no dibuje al cargar el formulario.
  iFormLoad = True

  PosForm Me

  For i = 1 To 63
    Load cell(i)
    cell(i).Visible = True
  Next i
    
  'Carga el cuadro de lista
  lstRedrawType.AddItem "Ninguno"
  lstRedrawType.AddItem "Usando variables"
  lstRedrawType.AddItem "Con cada imagen invisible"
  lstRedrawType.AddItem "Haciendo el contenedor invisible"
  lstRedrawType.ListIndex = 0

End Sub

Private Sub Form_Resize()
' No hace nada cuando está minimizado
    If WindowState <> 1 And iFormLoad <> True Then
        Container.Height = ScaleHeight - Container.Top
        ReDraw
    End If

    ' Restablece el indicador para volver a dibujar en la siguiente ocasión.
    iFormLoad = False

End Sub

Private Sub optGraphicType_Click(Index As Integer)
Dim i As Integer
    If Index Then
        cell(0).Picture = imgMetafile.Picture
    Else
        cell(0).Picture = imgBitmap.Picture
    End If
    Screen.MousePointer = 11
    ' Hace invisible el contenedor mientras se cargan imágenes.
    ' Haga comentarios estas líneas para ver las diferencias de velocidad.
    Container.Visible = False
    For i = 1 To 63
        cell(i).Picture = cell(0).Picture
    Next i
    Container.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub ReDraw()

Dim start, Finish
    
    Screen.MousePointer = 11
    start = Timer
    Select Case lstRedrawType.ListIndex
        Case 0
            ReSizeCells0
        Case 1
            ReSizeCells1
        Case 2
            ReSizeCells2
        Case 3
            ReSizeCells3
    End Select
    Finish = Timer
    lblTime = "Tiempo: " + Format(Finish - start, "0.000") + " Seg."
    Screen.MousePointer = 0

End Sub

Private Sub ReSizeCells0()
' Código nativo para cambiar el tamaño y volver a dibujar las celdas gráficas.
Dim i As Integer
    For i = 0 To 63
        cell(i).Move (i Mod 8) * (Container.ScaleWidth \ 8), (i \ 8) * (Container.ScaleHeight \ 8), Container.ScaleWidth \ 8, Container.ScaleHeight \ 8
    Next i
End Sub

Private Sub ReSizeCells1()
' Código mejorado: reemplaza las expresiones constantes por variables
' y utiliza variables en lugar de propiedades.
Dim i As Integer, cellwidth As Integer, cellheight As Integer
    cellwidth = Container.ScaleWidth \ 8
    cellheight = Container.ScaleHeight \ 8
    For i = 0 To 63
        cell(i).Move (i Mod 8) * cellwidth, (i \ 8) * cellheight, cellwidth, cellheight
    Next i
End Sub

Private Sub ReSizeCells2()
' Código más mejorado: hace invisible los controles Image
' al moverlos; esto aumenta la velocidad de dibujo.
Dim i As Integer, cellwidth As Integer, cellheight As Integer
    cellwidth = Container.ScaleWidth \ 8
    cellheight = Container.ScaleHeight \ 8
    For i = 0 To 63
        cell(i).Visible = False
        cell(i).Move (i Mod 8) * cellwidth, (i \ 8) * cellheight, cellwidth, cellheight
        cell(i).Visible = True
    Next i
End Sub

Private Sub ReSizeCells3()
' El código más rápido: hace invisible el contenedor durante el movimiento y el
' cambio de tamaño de todas las imágenes contenidas. Sólo se dibuja una vez.
Dim i As Integer, cellwidth As Integer, cellheight As Integer
    Container.Visible = False
    cellwidth = Container.ScaleWidth \ 8
    cellheight = Container.ScaleHeight \ 8
    For i = 0 To 63
        cell(i).Move (i Mod 8) * cellwidth, (i \ 8) * cellheight, cellwidth, cellheight
    Next i
    Container.Visible = True
    'Container.Refresh
End Sub

