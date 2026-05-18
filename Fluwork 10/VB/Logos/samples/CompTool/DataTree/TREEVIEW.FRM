VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmTreeview 
   Caption         =   "TreeView enlazado a base de datos"
   ClientHeight    =   5655
   ClientLeft      =   735
   ClientTop       =   2640
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   9660
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6960
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Cargar"
      Height          =   270
      Left            =   2190
      TabIndex        =   1
      Top             =   4935
      Visible         =   0   'False
      Width           =   1050
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   30
      Left            =   1305
      TabIndex        =   6
      Top             =   5145
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   327680
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "treeview.frx":0000
   End
   Begin ComctlLib.StatusBar sbrDB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5400
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "treeview.frx":001C
   End
   Begin ComctlLib.ProgressBar prgLoad 
      Height          =   210
      Left            =   255
      TabIndex        =   3
      Top             =   375
      Visible         =   0   'False
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   370
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "treeview.frx":0038
   End
   Begin ComctlLib.ListView lvwDB 
      Height          =   4215
      Left            =   3480
      TabIndex        =   2
      Top             =   615
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      Icons           =   "imlIcons"
      SmallIcons      =   "imlSmallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "treeview.frx":0054
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   1605
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FilterIndex     =   474
      FontSize        =   1.17491e-38
   End
   Begin ComctlLib.TreeView tvwDB 
      Height          =   4215
      Left            =   255
      TabIndex        =   0
      Top             =   615
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   7435
      _Version        =   327680
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "treeview.frx":0070
   End
   Begin VB.Label lblI 
      Caption         =   "Haga clic en Cargar, en el menú Archivo, para cargar Biblio.mdb"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   45
      Width           =   4695
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   645
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":008C
            Key             =   "book"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   45
      Top             =   4815
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":03A6
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":08C8
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":0DEA
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":130C
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":182E
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "treeview.frx":1E90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnuLoad 
         Caption         =   "Cargar"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmTreeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mDbBiblio As Database ' variable Database a nivel de módulo.
Private mNode As Node ' variable a nivel de módulo para nodos
Private mWksBiblio As Workspace ' variable WorkSpace a nivel de módulo.
Private mItem As ListItem ' variable ListItem a nivel de módulo.
Private EventFlag As Integer ' para señalar qué evento se ha producido.
Private mCurrentIndex As Integer ' indicador para asegurar que no se ha hecho ya clic en el nodo.
Private mStatusBarStyle As Integer ' Barra de estado de estilo interruptores.

Const PUBLISHER = 1 ' Para los objetos encabezado EventFlag, Signals Publisher.
Const TITLE = 2 ' EventFlag, señala Title en ListView
Private Sub cmdLoad_Click()
    ' Declara variables para los objetos de acceso a datos.
    Dim rsPublishers As Recordset
    Dim rsTitles As Recordset
    
    Dim intIndex ' Variable para el índice del nodo actual.
    
    ' Establece la base de datos BIBLIO.MDB en Database.
    ' Establece la tabla "Authors" en el conjunto de registros.
    Set rsPublishers = mDbBiblio. _
    OpenRecordset("Publishers", dbOpenDynaset)
    
    ' Muestra el control Progressbar.
    prgLoad.Visible = True
    ' Establece el estilo simple para StatusBar
    sbrDB.Style = sbrSimple
    ' Mientras el registro no sea el último, agrega un objeto ListItem.
    ' Utiliza el campo Name para el texto del objeto ListItem.
    Do Until rsPublishers.EOF
        
        prgLoad.Value = rsPublishers.PercentPosition
        ' Agrega un nodo al TreeView y establece sus propiedades.
        Set mNode = tvwDB.Nodes.Add(1, tvwChild, rsPublishers!PubID & " ID", CStr(rsPublishers!Name), "closed")
        mNode.Tag = "Publisher" ' Identifica la tabla.
        ' Establece en la variable intIndex la propiedad Index del
        ' objeto Node recién creado. Use esta variable para agregar
        ' objetos Node hijos al objeto Node actual.
        intIndex = mNode.Index
        ' Mientras está en este registro, crea un conjunto de
        ' registros con una consulta que sólo busca los títulos
        ' con el mismo PubID. Para cada registro del conjunto
        ' resultante, agrega un objeto Node al control TreeView
        ' y establece sus propiedades con los campos de título,
        ' ISBN y autor.
        Set rsTitles = mDbBiblio.OpenRecordset _
        ("select * from Titles Where PubID = " & rsPublishers!PubID)
        Do Until rsTitles.EOF
            Set mNode = tvwDB.Nodes.Add(intIndex, tvwChild)
            mNode.Text = rsTitles!TITLE ' Texto.
            mNode.Key = rsTitles!ISBN   ' ID único.
            mNode.Tag = "Authors"       ' Nombre de la tabla.
            mNode.Image = "smlBook"     ' Imagen de ImageList.
            rsTitles.MoveNext ' Pasa al siguiente registro de rsTitles.
        Loop
        rsPublishers.MoveNext   ' Pasa al siguiente registro de Publishers.
    Loop
    ' Oculta el control Progressbar
    prgLoad.Visible = False
    ' Establece el estilo normal para Statusbar.
    sbrDB.Style = sbrNormal
    ' Ordena los nodos Publishers.
    tvwDB.Nodes(1).Sorted = True
    ' Expande el nodo superior.
    tvwDB.Nodes(1).Expanded = True
    
    ' Configura la barra de estado.
    PublishersStatusBar
End Sub
Private Sub PublishersStatusBar()
    Dim rs As Recordset
    Set rs = mDbBiblio.OpenRecordset("Publishers", dbOpenTable)
    sbrDB.Panels(1).Text = "Nombre: " & rs.Name
    sbrDB.Panels(2).Text = "Fecha creación: " & _
    CStr(rs.DateCreated)
    sbrDB.Panels(3).Text = "Actualización: " & _
    rs.LastUpdated
    sbrDB.Panels(4).Text = "Bloqueado: " & rs.LockEdits
End Sub
Private Sub Combo1_Click()
    lvwDB.View = Combo1.ListIndex
End Sub


Private Sub Form_Load()
    ' Configura el control Combo1.
    With Combo1
        .AddItem "Iconos"      '0
        .AddItem "Iconos pequeños" '1
        .AddItem "Lista"         '2
        .AddItem "Informe"       '3
        .ListIndex = 3
    End With

    ' Configura el control ListView.
    lvwDB.View = lvwReport
        
    ' Configura el control ProgressBar.
    prgLoad.Max = 100
    
    ' Configura el control StatusBar.
    ' Agrega tres paneles, y establece Autosize para cada uno
    sbrDB.Panels.Add
    sbrDB.Panels.Add
    sbrDB.Panels.Add
    
    sbrDB.Panels(1).AutoSize = sbrContents
    sbrDB.Panels(2).AutoSize = sbrSpring
    sbrDB.Panels(3).AutoSize = sbrSpring
    sbrDB.Panels(4).AutoSize = sbrContents
    
    
    ' Configura cmdDialog por si no se encuentra Biblio.mdb.
    With dlgDialog
        .DialogTitle = "Imposible encontrar Biblio.mdb"
        
        .Filter = "(*.MDB)|*.mdb"
    End With
    
    On Error GoTo errFind
    ' Abre Biblio.MDB y establece la base de datos en la
    ' variable de objeto.
    Set mDbBiblio = DBEngine.OpenDatabase("BIBLIO.MDB")
    
    ' Configura el control TreeView
    tvwDB.Sorted = True
    Set mNode = tvwDB.Nodes.Add()
    'mNode.Text = mWksBiblio.Databases(0).Name
    mNode.Text = "Editores"
    mNode.Tag = mDbBiblio.Name
    mNode.Image = "closed"
    tvwDB.LabelEdit = False
   
    ' Si no se encuentra la base de datos Biblio, abre
    ' el control cuadro de diálogo común para que el usuario
    ' la busque.
errFind:
    If Err = 3024 Then
        dlgDialog.ShowOpen
        Set mDbBiblio = DBEngine.Workspaces(0).OpenDatabase(dlgDialog.filename)
        Resume Next
    ElseIf Err <> 0 Then ' otro error
        MsgBox "Error inesperado: " & Err.Description
        End
    ElseIf Err = 0 Then
        Exit Sub ' continúa
    End If
End Sub


Private Sub lvwDB_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    lvwDB.SortKey = ColumnHeader.Index - 1
    ' Establece Verdadero en Sorted para ordenar la lista.
    lvwDB.Sorted = True

End Sub

Private Sub lvwDB_ItemClick(ByVal Item As ListItem)
    GetData (Item.Key)
    
End Sub
Private Sub GetData(ISBN)
    If EventFlag <> TITLE Then
        sbrDB.Panels.Clear
        Dim pnlX As Panel
        Set pnlX = sbrDB.Panels.Add(, "ISBN")
        pnlX.AutoSize = sbrSpring
       Set pnlX = sbrDB.Panels.Add(, "author")
        pnlX.AutoSize = sbrSpring
        Set pnlX = sbrDB.Panels.Add(, "year")
        pnlX.Width = 1000
        Set pnlX = sbrDB.Panels.Add(, "description")
        pnlX.AutoSize = sbrSpring
    End If
    Dim rsTitles As Recordset
    ' Establece el conjunto de registros "Titles" en la variable de objeto
    Set rsTitles = mDbBiblio.OpenRecordset("Titles", dbOpenDynaset)
    
    ' Busca el número ISBN en la tabla Titles y rellena
    ' los paneles de StatusBar con información.
    Dim strQuery As String
    strQuery = "ISBN = " & "'" & ISBN & "'"
    rsTitles.FindFirst strQuery
    sbrDB.Panels("author").Text = GetAuthor(ISBN)
    sbrDB.Panels("ISBN").Text = rsTitles!ISBN
    If Not IsNull(rsTitles![Year Published]) Then
        sbrDB.Panels("year").Text = rsTitles![Year Published]
    Else
        sbrDB.Panels("year").Text = "n/a"
    End If
    If Not IsNull(rsTitles!Description) Then
        sbrDB.Panels("description").Text = rsTitles!Description
    Else
        sbrDB.Panels("description").Text = "n/a"
    End If
        
    
    ' Establece EventFlag para que no haya que volver
    ' a crear los paneles.
    EventFlag = TITLE
End Sub


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuLoad_Click()
    Static Loaded As Boolean
    If Loaded = True Then
        Exit Sub
    Else
        cmdLoad_Click
        Loaded = Abs(Loaded - 1)
        mnuLoad.Enabled = False
        
    End If
End Sub

Private Sub tvwDB_Collapse(ByVal Node As Node)
    If Node.Tag = "Publisher" Or Node.Index = 1 _
    Then Node.Image = "closed"
End Sub

Private Sub tvwDB_Expand(ByVal Node As Node)
    If Node.Tag = "Publisher" Or Node.Index = 1 Then
        Node.Image = "open"
        Node.Sorted = True
    End If
   
End Sub

Private Sub MakeColumns()
    ' Borra la colección ColumnHeaders.
    lvwDB.ColumnHeaders.Clear
    ' Agrega cuatro objetos ColumnHeaders.
    lvwDB.ColumnHeaders.Add , , "Título", 2000
    lvwDB.ColumnHeaders.Add , , "Autor"
    lvwDB.ColumnHeaders.Add , , "Año", 350
    lvwDB.ColumnHeaders.Add , , "ISBN"
    
    ' Establece la variable EventFlag para que
    ' esto no se haga una y otra vez.
    EventFlag = PUBLISHER
End Sub

Private Sub GetTitles(PubID)
    ' Muestra la barra de evolución
    prgLoad.Visible = True
    ' Borra los títulos anteriores
    lvwDB.ListItems.Clear
    ' Declara una variable de objeto de tipo Recordset
    Dim rsTitles As Recordset
    ' Mientras está en este registro, crea un conjunto de
    ' registros con una consulta que sólo busca los títulos con
    ' el mismo PubID. Para cada registro del resultado, agrega un
    ' objeto ListItem al control ListView, y establece las
    ' propiedades del nuevo objeto con los campos de título, ISBN
    ' y autor.
    Set rsTitles = mDbBiblio.OpenRecordset _
    ("select * from Titles where PubID = " & PubID)

    Do Until rsTitles.EOF
            prgLoad.Value = rsTitles.PercentPosition
            Set mItem = lvwDB.ListItems.Add _
            (, CStr(rsTitles!ISBN), rsTitles!TITLE, "book", "smlBook")
            
            ' Utiliza una función para obtener el autor y
            ' establece la propiedad SubItems(1).
            mItem.SubItems(1) = GetAuthor(rsTitles!ISBN)
            If Not IsNull(rsTitles![Year Published]) Then
                mItem.SubItems(2) = rsTitles![Year Published]
            End If
            mItem.SubItems(3) = rsTitles!ISBN
            rsTitles.MoveNext
    Loop
    prgLoad.Visible = False
    mCurrentIndex = PubID
End Sub

Private Function GetAuthor(ISBN)
    ' Devuelve el nombre de un autor de la tabla Titles.
    ' Abre dos conjuntos de registros. Busca "Title Author" para el ISBN,
    ' y luego usa el campo Au_ID para buscar el mismo valor en la tabla
    ' "Authors". Devuelve el nombre del autor.
    Dim rsTitlesAuthor As Recordset
    Dim rsAuthors As Recordset
    
    Set rsTitlesAuthor = mDbBiblio.OpenRecordset("Title Author", dbOpenDynaset)
    Set rsAuthors = mDbBiblio.OpenRecordset("Authors", dbOpenDynaset)
    
    Dim strQuery As String
    strQuery = "ISBN = " & "'" & ISBN & "'"
    rsTitlesAuthor.FindFirst strQuery
    If rsTitlesAuthor.NoMatch Then
        GetAuthor = "n/a"
        Exit Function
    Else
        strQuery = "Au_ID = " & rsTitlesAuthor!AU_ID
        rsAuthors.FindFirst strQuery
    End If
    GetAuthor = rsAuthors!Author
End Function

Private Sub tvwDB_NodeClick(ByVal Node As Node)
    ' Comprueba si Tag es "Publisher" y también la
    ' variable EventFlag para ver si ya se han creado
    ' ColumnHeaders. Si no es así, invoca el
    ' procedimiento MakeColumns.
    If Node.Tag = "Publisher" And EventFlag <> _
    PUBLISHER Then MakeColumns
    ' Si Tag es "Publisher" y el índice mItemCurrentIndex
    ' no coincide con la clave del nodo, invoca el
    ' procedimiento GetTitles.
    If Node.Tag = "Publisher" And mCurrentIndex <> Val(Node.Key) _
    Then GetTitles Val(Node.Key)
    If Node.Tag = "Publisher" Then
        sbrDB.Panels(1).Text = Node.Text
        sbrDB.Panels(2).Text = Node.Children & " títulos"
        Node.Sorted = True
    End If
End Sub
