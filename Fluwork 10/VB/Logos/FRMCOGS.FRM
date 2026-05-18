VERSION 5.00
Begin VB.Form frmCogs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de COGS"
   ClientHeight    =   2550
   ClientLeft      =   1215
   ClientTop       =   1515
   ClientWidth     =   7725
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   7725
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "&Guardar"
      Default         =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   6420
      TabIndex        =   13
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ay&uda"
      Height          =   420
      Index           =   1
      Left            =   6420
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   420
      Index           =   2
      Left            =   6420
      TabIndex        =   11
      Top             =   180
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de papel"
      Height          =   2340
      Left            =   2220
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton optPaperGrade 
         Caption         =   "Grado 2"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   690
         Width           =   975
      End
      Begin VB.OptionButton optPaperGrade 
         Caption         =   "Grado 3"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   1050
         Width           =   975
      End
      Begin VB.OptionButton optPaperGrade 
         Caption         =   "Grado 4"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   1410
         Width           =   975
      End
      Begin VB.OptionButton optPaperGrade 
         Caption         =   "Grado 1"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colores de la imagen"
      Height          =   1125
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2010
      Begin VB.OptionButton optPicColor 
         Caption         =   "Color"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   690
         Width           =   1500
      End
      Begin VB.OptionButton optPicColor 
         Caption         =   "Blanco y negro"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   345
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de encuadernación"
      Height          =   1125
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   2010
      Begin VB.OptionButton optBinding 
         Caption         =   "Tapa blanda"
         Height          =   285
         Index           =   1
         Left            =   140
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optBinding 
         Caption         =   "Tapa dura"
         Height          =   285
         Index           =   0
         Left            =   140
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame FrameConverted 
      BackColor       =   &H00C0C0C0&
      Height          =   2355
      Left            =   3540
      TabIndex        =   14
      Top             =   120
      Width           =   2805
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   2625
         TabIndex        =   32
         Top             =   1455
         Width           =   90
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   1980
         TabIndex        =   31
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo unitario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   1755
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_____________________________"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   1500
         Width           =   2610
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   1980
         TabIndex        =   28
         Top             =   1455
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de encuadernación"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   1455
         Width           =   1785
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   1980
         TabIndex        =   26
         Top             =   1245
         Width           =   585
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   2010
         TabIndex        =   25
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   24
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   23
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   22
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_____________________________"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   1020
         Width           =   2610
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   2625
         TabIndex        =   20
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Páginas"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   2625
         TabIndex        =   17
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grado del papel"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Colores de la imagen"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_____________________________"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   540
         Width           =   2610
      End
   End
End
Attribute VB_Name = "frmCogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRetVal As Long

' Estas variables contienen los valores de COGS actuales
' que ha seleccionado el usuario.
Dim cBindingCost As Currency
Dim cPicture As Currency
Dim cPaperGrade As Currency
Dim lngNumPages As Long

Dim strBinding As String
Dim strPicture As String
Dim strPaper As String

Dim acurBindingCost(1) As Currency
Dim acurPictureCost(1) As Currency
Dim acurPaperCost(3) As Currency
Dim acurCogs(5) As Currency

Private Sub optPicColor_Click(Index As Integer)
If Index = 0 Then
  ' Blanco y negro
  strPicture = "BN"
Else
  ' Color
  strPicture = "Color"
End If
  
acurCogs(0) = acurPictureCost(Index)
lblCost(0).Caption = Format$(acurCogs(0), "#.0000")

CalcUnitCost

End Sub

Private Sub Command1_Click(Index As Integer)
  Dim sHelpString As String
  
  Select Case Index
    Case 0  ' Guardar
      frmRevenue.txtRevParm(0).Text = strBinding & ", " & strPicture & ", " & strPaper
      frmRevenue.txtRevParm(0).Tag = lblCost(6).Caption
      Unload Me
      
    Case 1  ' Ayuda
      sHelpString = "Seleccione el tipo de encuadernación deseado, el tipo de imagen y la calidad del papel." & vbCrLf & _
                    "Haga clic en Guardar para guardar estos valores y volver al formulario Ingresos por libros." & vbCrLf & _
                    "Haga clic en Cancelar para descartar los valores y volver al formulario Ingresos por libros."
                    
      lRetVal = MsgBox(prompt:=sHelpString, Title:="Ayuda de COGS", Buttons:=vbInformation)
      
    Case 2  ' Cerrar
      Unload Me
      
 End Select
      
End Sub
Sub CalcUnitCost()
  Dim curTemp As Currency
    
  curTemp = acurCogs(0) + acurCogs(1)
  lblCost(2).Caption = Format$(curTemp, "#.0000")
  
  curTemp = curTemp * lngNumPages
  lblCost(4).Caption = Format$(curTemp, "#.0000")
  
  curTemp = curTemp + acurCogs(5)
  lblCost(6).Caption = Format$(curTemp, "#.0000")
  
End Sub
Function GetCOGS() As Integer
  Dim fld As Field
  Dim strSQL As String
  
  On Error GoTo GetCOGSError
  
  goStatusPanel.Text = "Buscando números COGS..."
  strSQL = "SELECT * FROM COGS"
           
  Set gSn = gDB.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
  
  acurBindingCost(0) = gSn.Fields("HardCover")
  acurBindingCost(1) = gSn.Fields("PaperBack")
  acurPictureCost(0) = gSn.Fields("BlackWhite")
  acurPictureCost(1) = gSn.Fields("Color")
  acurPaperCost(0) = gSn.Fields("Grade1")
  acurPaperCost(1) = gSn.Fields("Grade2")
  acurPaperCost(2) = gSn.Fields("Grade3")
  acurPaperCost(3) = gSn.Fields("Grade4")

  gSn.Close
  
  GetCOGS = True
  Exit Function
  
GetCOGSError:
  If IsObject(gSn) Then gSn.Close
  GetCOGS = False
    
End Function

Public Function GetBookPages(strTitle As String) As Long
Dim strSQL As String

Static strOldTitle As String
Static lngPages As Long

On Error GoTo GetBookPagesError

If strTitle <> strOldTitle Then
  goStatusPanel.Text = "Obteniendo el número de páginas de " & strTitle & "..."
  strSQL = "SELECT Titles.Pages " & _
           "FROM Titles " & _
           "WHERE ((Titles.Title=" & Chr$(34) & Trim$(strTitle) & Chr$(34) & "));"
  
  Set gSn = gDB.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
  lngPages = gSn.Fields("Pages")
  strOldTitle = strTitle
  gSn.Close
  
Else
  goStatusPanel.Text = "Usando el número de páginas almacenado."

End If

GetBookPages = lngPages

Exit Function

GetBookPagesError:
  If IsObject(gSn) Then gSn.Close
  GetBookPages = False
  MsgBox Error$ & Str$(Err), vbCritical, "Error en GetBookPages"

End Function

Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Show
  DoEvents
  
  lngNumPages = GetBookPages(frmRevenue.cboBooks.Text)
  lRetVal = GetCOGS
  
  lblCost(0).Caption = Format$(acurPictureCost(0), "#.0000")
  acurCogs(0) = acurPictureCost(0)
  
  lblCost(1).Caption = Format$(acurPaperCost(0), "#.0000")
  acurCogs(1) = acurPaperCost(0)
  
  lblCost(3).Caption = Str$(lngNumPages)
  
  lblCost(5).Caption = Format$(acurBindingCost(0), "#.0000")
  acurCogs(5) = acurBindingCost(0)
  
  strBinding = "Dura"
  strPicture = "BN"
  strPaper = "1"
  CalcUnitCost
  DoEvents
  
End Sub

Private Sub optBinding_Click(Index As Integer)
  If Index = 0 Then
    strBinding = "Dura"
    
  Else
    strBinding = "Blanda"
    
  End If
  
  acurCogs(5) = acurBindingCost(Index)
  lblCost(5).Caption = Format$(acurCogs(5), "#.0000")
  
  If optPaperGrade(0).Value Then
    optPaperGrade_Click (0)
  
  ElseIf optPaperGrade(0).Value Then
    optPaperGrade_Click (1)
    
  ElseIf optPaperGrade(2).Value Then
    optPaperGrade_Click (2)
    
  Else
    optPaperGrade_Click (3)
    
  End If

  CalcUnitCost
  
End Sub

Private Sub optPaperGrade_Click(Index As Integer)
  
  Select Case Index
    Case 0
      strPaper = "1"
      
    Case 1
      strPaper = "2"
      
    Case 2
      strPaper = "3"

    Case 3
      strPaper = "4"
      
  End Select
 
  
  If optBinding(1).Value = True Then
    acurCogs(1) = acurPaperCost(Index) * 0.5
      
  Else
    acurCogs(1) = acurPaperCost(Index)
    
  End If
  
  lblCost(1).Caption = Format$(acurCogs(1), "#.0000")
  CalcUnitCost
  
End Sub


