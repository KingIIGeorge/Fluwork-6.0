VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Examinador de base de datos bibliográficas"
   ClientHeight    =   6840
   ClientLeft      =   390
   ClientTop       =   1320
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   PaletteMode     =   1   'UseZOrder
   ScaleHeight     =   6840
   ScaleWidth      =   9000
   StartUpPosition =   2   'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox DisplayGroup 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Index           =   2
      Left            =   1080
      ScaleHeight     =   5235
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   1320
      Width           =   8055
      Begin VB.ListBox lstPubsTitles 
         Height          =   1260
         Left            =   3120
         TabIndex        =   45
         Top             =   2880
         Width           =   4575
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Comments"
         DataSource      =   "dbCtrl(2)"
         Height          =   975
         Index           =   9
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Fax"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Text            =   "Fax"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Telephone"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Text            =   "Telephone"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Zip"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   6
         Left            =   6120
         TabIndex        =   7
         Text            =   "Zip"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "State"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   5
         Left            =   5040
         TabIndex        =   6
         Text            =   "State"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "City"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Text            =   "City"
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Company Name"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Text            =   "Company Name"
         Top             =   1080
         Width           =   7335
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Name"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Text            =   "Name"
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "PubID"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Text            =   "Name"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Data dbCtrl 
         BackColor       =   &H00FFFF00&
         Caption         =   "  Editores"
         Connect         =   "Access"
         DatabaseName    =   "C:\VB5\BIBLIO.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3000
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Publishers"
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox txtPubs 
         DataField       =   "Address"
         DataSource      =   "dbCtrl(2)"
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Text            =   "Address"
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Títulos publicados"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   44
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentarios"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   43
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   42
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   41
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código postal"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   6120
         TabIndex        =   40
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   39
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ciudad"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   38
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dirección"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre de la compañía"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   36
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblPubs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Editores"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   34
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox DisplayGroup 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Index           =   1
      Left            =   480
      ScaleHeight     =   5235
      ScaleWidth      =   7995
      TabIndex        =   11
      Top             =   900
      Width           =   8055
      Begin VB.ListBox lstAuthorTitles 
         Height          =   870
         Left            =   2880
         TabIndex        =   53
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtAuthors 
         DataField       =   "Year Born"
         DataSource      =   "dbCtrl(1)"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   49
         Text            =   "Year Born"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtAuthors 
         DataField       =   "Au_ID"
         DataSource      =   "dbCtrl(1)"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   48
         Text            =   "Au_ID"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtAuthors 
         DataField       =   "Author"
         DataSource      =   "dbCtrl(1)"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Text            =   "Author"
         Top             =   600
         Width           =   6975
      End
      Begin VB.Data dbCtrl 
         BackColor       =   &H00FFFF00&
         Caption         =   "    Autores"
         Connect         =   "Access"
         DatabaseName    =   "C:\VB5\BIBLIO.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3000
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Authors"
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label lblAuthors 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Títulos escritos"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   52
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblAuthors 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año de nacimiento"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   51
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lblAuthors 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Id. del autor"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   50
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblAuthors 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Autores"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   47
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Data dcAuthorTitles 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB5\BIBLIO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dcTitleAuthors 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB5\BIBLIO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   0
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dcPubsTitles 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB5\BIBLIO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   0
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9075
      TabIndex        =   13
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton btnGroup 
         Caption         =   "Editores"
         Height          =   495
         Index           =   2
         Left            =   6120
         TabIndex        =   16
         Top             =   0
         Width           =   3015
      End
      Begin VB.CommandButton btnGroup 
         Caption         =   "Autores"
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   0
         Width           =   3135
      End
      Begin VB.CommandButton btnGroup 
         Caption         =   "Títulos"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      FilterIndex     =   639
      FontSize        =   1.17491e-38
      HelpCommand     =   3
      HelpFile        =   "BIBLIO.HLP"
   End
   Begin VB.PictureBox DisplayGroup 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Index           =   0
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   7995
      TabIndex        =   17
      Top             =   480
      Width           =   8055
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PubID"
         DataSource      =   "dbCtrl(0)"
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Text            =   "PubID"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Comments"
         DataSource      =   "dbCtrl(0)"
         Height          =   1095
         Index           =   5
         Left            =   4080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Subject"
         DataSource      =   "dbCtrl(0)"
         Height          =   1095
         Index           =   4
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Description"
         DataSource      =   "dbCtrl(0)"
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Text            =   "Description"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ISBN"
         DataSource      =   "dbCtrl(0)"
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Text            =   "ISBN"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Year Published"
         DataSource      =   "dbCtrl(0)"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Text            =   "Year"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtTitles 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Title"
         DataSource      =   "dbCtrl(0)"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Text            =   "Título"
         Top             =   480
         Width           =   7215
      End
      Begin VB.Data dbCtrl 
         BackColor       =   &H00FFFF00&
         Caption         =   "      Títulos"
         Connect         =   "Access"
         DatabaseName    =   "C:\VB5\BIBLIO.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2880
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "Titles"
         Top             =   4800
         Width           =   2175
      End
      Begin VB.ListBox lstTitleAuthors 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   4080
         TabIndex        =   46
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblTitlePub 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   1920
         TabIndex        =   54
         Top             =   3720
         Width           =   5535
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Autores"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   4080
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Editores"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   32
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Id. del editor"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentarios"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   30
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tema"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ISBN"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Títulos"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "&Selección"
      Begin VB.Menu mnuTitles 
         Caption         =   "&Títulos"
      End
      Begin VB.Menu mnuAuthors 
         Caption         =   "&Autores"
      End
      Begin VB.Menu mnuPublishers 
         Caption         =   "&Editores"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaración de la API de Windows para invocar WinHelp.
'Declaraciones de constantes para las matrices de controles Data y cuadro de imagen
Const Titles% = 0
Const Authors% = 1
Const Publishers% = 2

'Almacena el directorio de Biblio.Mdb para su referencia dentro del módulo
Dim sDir As String

'Enumera todos los títulos de un editor concreto
Sub ListPubsTitles()
    Dim Qry As String
    'Genera la consulta
    Qry = "SELECT DISTINCTROW Titles.Title FROM Titles"
    Qry = Qry + " WHERE (Titles.PubID=" + txtPubs(0).Text + ")"
    'Actualiza el control Data
    dcPubsTitles.RecordSource = Qry
    dcPubsTitles.Refresh
    'Borra el cuadro de lista y genera la nueva lista
    lstPubsTitles.Clear
    If dcPubsTitles.Recordset.EOF Then
        lstPubsTitles.AddItem "El editor no tiene títulos."
    End If
    Do While Not dcPubsTitles.Recordset.EOF
        lstPubsTitles.AddItem dcPubsTitles.Recordset(0)
        dcPubsTitles.Recordset.MoveNext
    Loop
End Sub

'Enumera todos los títulos de un autor concreto
Sub ListAuthorTitles()
    Dim Qry As String
    'Genera la consulta
    Qry = "SELECT Titles.Title FROM Titles"
    Qry = Qry + " INNER JOIN [Title Author] ON"
    Qry = Qry + " Titles.ISBN = [Title Author].ISBN WHERE"
    Qry = Qry + " (([Title Author].Au_ID=" + txtAuthors(1).Text + "))"
    'Actualiza el control Data
    dcAuthorTitles.RecordSource = Qry
    dcAuthorTitles.Refresh
    'Borra el cuadro de lista y genera la nueva lista
    lstAuthorTitles.Clear
    Do While Not dcAuthorTitles.Recordset.EOF
        lstAuthorTitles.AddItem dcAuthorTitles.Recordset(0)
        dcAuthorTitles.Recordset.MoveNext
    Loop
End Sub

'Enumera los autores de un título concreto
Sub ListTitleAuthors()
    Dim Qry As String
    'Genera la consulta
    Qry = "SELECT Authors.Author FROM Authors"
    Qry = Qry + " INNER JOIN [Title Author] ON"
    Qry = Qry + " Authors.Au_ID = [Title Author].Au_ID WHERE"
    Qry = Qry + " (([Title Author].ISBN='" + txtTitles(2).Text + "'))"
    'Actualiza el control Data
    dcTitleAuthors.RecordSource = Qry
    dcTitleAuthors.Refresh
    'Borra el cuadro de lista y genera la nueva lista
    lstTitleAuthors.Clear
    Do While Not dcTitleAuthors.Recordset.EOF
        lstTitleAuthors.AddItem dcTitleAuthors.Recordset(0)
        dcTitleAuthors.Recordset.MoveNext
    Loop
    
    'Obtiene el nombre del editor a partir del texto de PubID
    Dim myDB As Database, myRS As Recordset
    'Establece la base de datos a BIBLIO.MDB
    Set myDB = DBEngine.Workspaces(0).OpenDatabase(sDir & "\biblio.mdb")
    'Establece el conjunto de registros a la tabla Publishers.
    Set myRS = myDB.OpenRecordset("Publishers", dbOpenDynaset)
    Dim myCriteria As String
    myCriteria = "PubID = " & txtTitles(6).Text
    myRS.FindFirst myCriteria
    lblTitlePub.Caption = myRS!Name
    
End Sub

'Pone en primer plano el cuadro de imagen seleccionado haciendo clic en el botón
Private Sub btnGroup_Click(Index As Integer)
    DisplayGroup(Index).ZOrder 0
End Sub

'Genera las listas de registros relacionados cuando cambia el registro actual
Private Sub dbCtrl_Reposition(Index As Integer)
    If Index = Publishers Then
        ListPubsTitles
    ElseIf Index = Authors Then
        ListAuthorTitles
    ElseIf Index = Titles Then
        ListTitleAuthors
    End If
End Sub

'Pone en primer plano el cuadro de imagen seleccionado haciendo clic en el botón
Private Sub DisplayGroup_Click(Index As Integer)
    DisplayGroup(Index).ZOrder 0
End Sub

Private Sub Form_Load()
  On Error GoTo FormLoad_Err
 
  'Establece la variable de cadena de texto a nivel de módulo
  sDir = CurDir
  
  dcAuthorTitles.DatabaseName = sDir & "\biblio.mdb"
  dcAuthorTitles.Refresh
  dcPubsTitles.DatabaseName = sDir & "\biblio.mdb"
  dcPubsTitles.Refresh
  dcTitleAuthors.DatabaseName = sDir & "\biblio.mdb"
  dcTitleAuthors.Refresh
  dbCtrl(0).DatabaseName = sDir & "\biblio.mdb"
  dbCtrl(0).Refresh
  dbCtrl(1).DatabaseName = sDir & "\biblio.mdb"
  dbCtrl(1).Refresh
  dbCtrl(2).DatabaseName = sDir & "\biblio.mdb"
  dbCtrl(2).Refresh
  
FormLoad_Exit:
Exit Sub
  
FormLoad_Err:
    'Se requiere un control de errores avanzado si no se encuentra la base de datos,
    'como indicaría el error 3024.
    If Err = 3024 Then
        'Establece las propiedades de CommonDialog antes de mostrarlo
        With CommonDialog1
          .DialogTitle = "Imposible encontrar la ubicación del archivo biblio.mdb"
          .Filter = "(*.mdb)|*.mdb"
          .InitDir = CurDir
          .filename = "biblio.mdb"
          .Flags = cdlOFNReadOnly Or cdlOFNExplorer Or _
           cdlOFNFileMustExist Or cdlOFNPathMustExist
          .ShowOpen
          'Asegura que el nombre del archivo no es una cadena vacía
          If .filename <> "" Then
            'Asegura que el archivo de base de datos devuelto es realmente biblio.mdb
            If Right(UCase(.filename), Len("\biblio.mdb")) = "\BIBLIO.MDB" Then
            dcAuthorTitles.DatabaseName = .filename
            'Establece la variable global sDir a la ruta devuelta por Filename en CommonDialog
            sDir = Mid(.filename, 1, Len(.filename) - Len("\biblio.mdb"))
            End If
            Resume
          Else
            Unload Me
          End If
        End With
    ElseIf Err <> 0 Then ' otro error
        MsgBox "Error inesperado: " & Err.Description
        End
    End If
      
    Resume FormLoad_Exit
  
End Sub

'Pasa al editor seleccionado cuando se hace clic en el cuadro de texto Editores
'del formulario Títulos
Private Sub lblTitlePub_Click()
    Dim BM As Variant, FindStr As String
    FindStr = "PubID = " + txtTitles(6).Text
    Debug.Print FindStr
    BM = dbCtrl(Publishers).Recordset.Bookmark
    dbCtrl(Publishers).Recordset.FindFirst FindStr
     If dbCtrl(Publishers).Recordset.NoMatch = True Then
        dbCtrl(Publishers).Recordset.Bookmark = BM
    Else
        DisplayGroup_Click (Publishers)
    End If
    
End Sub

'Pasa al título seleccionado cuando se hace clic en el cuadro de lista Títulos escritos
'del formulario Autores
Private Sub lstAuthorTitles_Click()
    Dim BM As Variant, FindStr As String
    FindStr = "Title = """ + lstAuthorTitles.List(lstAuthorTitles.ListIndex) + """"
    BM = dbCtrl(Titles).Recordset.Bookmark
    dbCtrl(Titles).Recordset.FindFirst FindStr
    If dbCtrl(Titles).Recordset.NoMatch = True Then
        dbCtrl(Titles).Recordset.Bookmark = BM
    Else
        DisplayGroup_Click (Titles)
    End If
End Sub

'Pasa al título seleccionado cuando se hace clic en el cuadro de lista Títulos publicados
'del formulario Editores
Private Sub lstPubsTitles_Click()
    Dim BM As Variant
    BM = dbCtrl(Titles).Recordset.Bookmark
    dbCtrl(Titles).Recordset.FindFirst _
        "Title = """ + lstPubsTitles.List(lstPubsTitles.ListIndex) + """"
    If dbCtrl(Titles).Recordset.NoMatch = True Then
        dbCtrl(Titles).Recordset.Bookmark = BM
    Else
        DisplayGroup_Click (Titles)
    End If
End Sub

'Pasa al autor seleccionado cuando se hace clic en el cuadro de lista Autores
'del formulario Títulos
Private Sub lstTitleAuthors_Click()
    Dim BM As Variant
    BM = dbCtrl(Authors).Recordset.Bookmark
    dbCtrl(Authors).Recordset.FindFirst _
        "Author = """ + lstTitleAuthors.List(lstTitleAuthors.ListIndex) + """"
    If dbCtrl(Authors).Recordset.NoMatch = True Then
        dbCtrl(Authors).Recordset.Bookmark = BM
    Else
        DisplayGroup_Click (Authors)
    End If
End Sub

'Pone el cuadro de imagen del autor en primer plano
Private Sub mnuAuthors_Click()
    DisplayGroup_Click (Authors)
End Sub

'Pone el cuadro de imagen del editor en primer plano
Private Sub mnuPublishers_Click()
    DisplayGroup_Click (Publishers)
End Sub

'Pone el cuadro de imagen de títulos en primer plano
Private Sub mnuTitles_Click()
    DisplayGroup_Click (Titles)
End Sub

'Termina el programa
Private Sub mnuExit_Click()
    Unload Me
End Sub

'Presenta el cuadro Acerca de
Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1
End Sub

'Abre la Ayuda y presenta la pantalla de contenido
Private Sub mnuHelpContents_Click()
  With CommonDialog1
    .HelpCommand = cdlHelpContents
    .HelpFile = "C:\vb\samples\datactrl\biblio.hlp"
    .ShowHelp
    End With
End Sub




