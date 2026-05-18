VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmRevenue 
   Caption         =   "Ingresos por libros"
   ClientHeight    =   5820
   ClientLeft      =   570
   ClientTop       =   600
   ClientWidth     =   6795
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BOOK_CLI.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   6795
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de ingresos"
      Height          =   2685
      Left            =   80
      TabIndex        =   4
      Top             =   2220
      Width           =   6660
      Begin VB.CheckBox chkDiscount 
         Caption         =   "Descuento promocional"
         Height          =   270
         Left            =   4080
         TabIndex        =   8
         Top             =   960
         Width           =   2115
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   3915
         TabIndex        =   15
         Top             =   1020
         Width           =   2595
         Begin VB.TextBox txtRevParm 
            Enabled         =   0   'False
            Height          =   300
            Index           =   5
            Left            =   1755
            TabIndex        =   10
            Text            =   "3"
            Top             =   690
            Width           =   375
         End
         Begin VB.TextBox txtRevParm 
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   1755
            TabIndex        =   9
            Text            =   "2"
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox txtRevParm 
            Enabled         =   0   'False
            Height          =   300
            Index           =   6
            Left            =   1755
            TabIndex        =   11
            Text            =   "15"
            Top             =   1080
            Width           =   375
         End
         Begin ComCtl2.UpDown udDiscount 
            Height          =   255
            Index           =   4
            Left            =   2280
            TabIndex        =   32
            Top             =   360
            Width           =   150
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
         End
         Begin ComCtl2.UpDown udDiscount 
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   33
            Top             =   720
            Width           =   150
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
         End
         Begin ComCtl2.UpDown udDiscount 
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   34
            Top             =   1080
            Width           =   150
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración (meses):"
            Enabled         =   0   'False
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   25
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes de inicio:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   24
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento %:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   555
            TabIndex        =   23
            Top             =   1080
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdCogs 
         Caption         =   "..."
         Height          =   330
         Left            =   3720
         TabIndex        =   3
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtRevParm 
         Height          =   300
         Index           =   3
         Left            =   2445
         TabIndex        =   7
         Text            =   "5000"
         Top             =   2145
         Width           =   1215
      End
      Begin VB.TextBox txtRevParm 
         Height          =   300
         Index           =   2
         Left            =   2445
         TabIndex        =   6
         Text            =   "12"
         Top             =   1590
         Width           =   1215
      End
      Begin VB.TextBox txtRevParm 
         Height          =   300
         Index           =   1
         Left            =   2445
         TabIndex        =   5
         Text            =   "10000"
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox txtRevParm 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   405
         Width           =   2160
      End
      Begin VB.TextBox txtRevParm 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   4215
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1245
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades/mes:"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   22
         Top             =   2190
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo (meses):"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   21
         Top             =   1620
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Costo de publicidad:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   1050
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código COGS:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   19
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Derechos de autor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   4320
         TabIndex        =   27
         Top             =   1350
         Width           =   1080
      End
   End
   Begin VB.Frame frmSalesModel 
      Caption         =   "Modelo de ventas"
      Height          =   1350
      Left            =   80
      TabIndex        =   2
      Top             =   765
      Width           =   6660
      Begin VB.OptionButton optAnalysis 
         Height          =   975
         Index           =   2
         Left            =   4920
         Picture         =   "BOOK_CLI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optAnalysis 
         Height          =   975
         Index           =   1
         Left            =   2580
         Picture         =   "BOOK_CLI.frx":166C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optAnalysis 
         Height          =   975
         Index           =   0
         Left            =   240
         Picture         =   "BOOK_CLI.frx":29CE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   420
      Left            =   5520
      TabIndex        =   14
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Ay&uda"
      Height          =   420
      Left            =   80
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Ejecutar..."
      Default         =   -1  'True
      Height          =   420
      Left            =   4080
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox cboBooks 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4620
   End
   Begin VB.ComboBox cboAuthors 
      Height          =   315
      Left            =   80
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1980
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   26
      Top             =   5490
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FontSize        =   1.17485e-38
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Libros:"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autor:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmRevenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String

Private Sub cboAuthors_Click()
  
  GetTitles cboAuthors.Text

End Sub

Sub GetAuthors()
  Dim strSQL As String
  Dim objAuthorField As Object
  
  On Error GoTo GetAuthorsError

  strSQL = "SELECT Authors.Author " & _
                "FROM Authors"

  Set gSn = gDB.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
  
  Set objAuthorField = gSn.Fields("author")
  
  Do While Not gSn.EOF
    cboAuthors.AddItem objAuthorField
    gSn.MoveNext
  
  Loop
  
  gSn.Close
  cboAuthors.ListIndex = 0

  Exit Sub

GetAuthorsError:
    If IsObject(gSn) Then gSn.Close
    Exit Sub
  
End Sub
Sub GetTitles(strAuthor As String)
Dim strSQL As String
Dim objTitlesField As Object

On Error GoTo GetTitlesError

strSQL = "SELECT Titles.Title " & _
         "FROM Titles INNER JOIN (Authors INNER JOIN [Title Author] ON Authors.Au_ID = [Title Author].Au_ID) ON Titles.ISBN = [Title Author].ISBN " & _
         "WHERE ((Authors.Author=" & Chr$(34) & strAuthor & Chr$(34) & "));"

Set gSn = gDB.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
  
cboBooks.Clear

Set objTitlesField = gSn.Fields("Title")

Do While Not gSn.EOF
  cboBooks.AddItem objTitlesField
  gSn.MoveNext
  
Loop

gSn.Close
cboBooks.ListIndex = 0

Exit Sub

GetTitlesError:
  If Not gSn Is Nothing Then gSn.Close
  Exit Sub
  
End Sub

Private Sub chkDiscount_Click()
    Label2(4).Enabled = Not Label2(4).Enabled
    Label2(5).Enabled = Not Label2(5).Enabled
    Label2(6).Enabled = Not Label2(6).Enabled
    txtRevParm(4).Enabled = Not txtRevParm(4).Enabled
    txtRevParm(5).Enabled = Not txtRevParm(5).Enabled
    txtRevParm(6).Enabled = Not txtRevParm(6).Enabled
    udDiscount(4).Enabled = Not udDiscount(4).Enabled
    udDiscount(5).Enabled = Not udDiscount(5).Enabled
    udDiscount(6).Enabled = Not udDiscount(6).Enabled
    
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdCogs_Click()
  goStatusPanel.Text = "Calculando el costo de los artículos."
  Load frmCogs
  goStatusPanel.Text = ""
  
End Sub

Private Sub cmdHelp_Click()
Dim sHelpString As String
      
    sHelpString = "1. Seleccione un autor y el título de un libro." & vbCrLf & _
                  "2. Seleccione un modelo de ventas." & vbCrLf & _
                  "3. Haga clic en el botón de selección para establecer el precio" & vbCrLf & _
                  "    COGS." & vbCrLf & _
                  "4. Escriba una cantidad para gastar en costo de publicidad." & vbCrLf & _
                  "5. Escriba el número de meses del periodo de ventas." & vbCrLf & _
                  "6. Escriba el número de unidades vendidas por mes." & vbCrLf & _
                  "7. Puede hacer clic en la casilla de Descuento promocional" & vbCrLf & _
                  "    para establecer un descuento de ventas para este libro." & vbCrLf & _
                  "8. Haga clic en el botón Ejecutar para ver un gráfico de" & vbCrLf & _
                  "    los beneficios del editor y los derechos de autor."

    MsgBox sHelpString, vbInformation, "Ayuda de ingresos por libros"
End Sub

Private Sub cmdExecute_Click()

  If cboBooks.Text = "<None>" Or cboBooks.Text = "" Then
    MsgBox "Imposible crear un gráfico si no hay ningún título disponible.", _
                  vbExclamation, "Error al crear el gráfico"
    Exit Sub
  End If
                         
  goStatusPanel.Text = "Creando el gráfico"
  Load frmChart
  frmChart.Show
      
End Sub

Private Sub Form_Load()
  
Screen.MousePointer = vbHourglass

Me.Show

On Error GoTo LoadError

  Set goStatusPanel = StatusBar1.Panels.Item(1)
  Set gDB = Workspaces(0).OpenDatabase(gstrDBName, False, True)
  
  goStatusPanel.Text = "Obteniendo la lista de autores..."
  GetAuthors
  
  goStatusPanel.Text = "Conectándose con el dominio remoto..."
  Set goBookSaleSvrInst = CreateObject("BookSaleSvr.Sales")
  goStatusPanel.Text = "Conexión remota terminada satisfactoriamente..."
  
LoadExit:
Screen.MousePointer = vbDefault
Exit Sub
  
LoadError:
  
  'Es necesario un tratamiento de errores avanzado si no se encuentra
  'la base de datos, como indicaría el error 3024.
  If Err = 3024 Then
      'Establece las propiedades de CommonDialog antes de mostrarlo
      With CommonDialog1
        .DialogTitle = "Imposible encontrar la ubicación del archivo booksale.mdb"
        .Filter = "(*.mdb)|*.mdb"
        .InitDir = CurDir
        .filename = ""
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
        .ShowOpen
        'Se asegura de que el nombre del archivo no sea una cadena vacía
        If .filename <> "" Then
          'Se asegura de que el archivo de base de datos obtenido es realmente
          'booksale.mdb
          If Right(UCase(.filename), Len("booksale.mdb")) = "BOOKSALE.MDB" Then
            gstrDBName = .filename
          End If
          Resume
        Else
          Unload Me
        End If
      End With
  ElseIf Err <> 0 Then ' otro error
      MsgBox Error$ & " - " & Str$(Err), vbExclamation, "Error en Form_Load"
      End
  End If
  
  Resume LoadExit

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set gSn = Nothing
  Set goBookSaleSvrInst = Nothing

End Sub

Private Sub optAnalysis_Click(Index As Integer)

  frmSalesModel.Tag = Index

End Sub

Private Sub txtRevParm_GotFocus(Index As Integer)
  txtRevParm(Index).SelStart = 0
  txtRevParm(Index).SelLength = Len(txtRevParm(Index).Text)
  
End Sub

Private Sub udDiscount_DownClick(Index As Integer)

  If Val(txtRevParm(Index).Text) > 1 Then
    txtRevParm(Index).Text = Val(txtRevParm(Index).Text) - 1
  End If

End Sub

Private Sub udDiscount_UpClick(Index As Integer)

  If Val(txtRevParm(Index).Text) < 99 Then
    txtRevParm(Index).Text = Val(txtRevParm(Index).Text) + 1
  End If

End Sub

