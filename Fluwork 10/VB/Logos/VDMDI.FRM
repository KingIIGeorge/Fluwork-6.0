VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "VisData"
   ClientHeight    =   6780
   ClientLeft      =   4110
   ClientTop       =   2625
   ClientWidth     =   10005
   HelpContextID   =   2016116
   Icon            =   "VDMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin ComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      HelpContextID   =   65278
      HelpFile        =   $"VDMDI.frx":014A
      ImageList       =   "imlToolbarPics"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Table"
            Object.ToolTipText     =   "Recordset de tipo Table"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Dynaset"
            Object.ToolTipText     =   "Recordset de tipo Dynaset"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Snapshot"
            Object.ToolTipText     =   "Recordset de tipo Snapshot"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "PassThrough"
            Object.ToolTipText     =   "Recordset de tipo Passthrough"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DataControl"
            Object.ToolTipText     =   "Usar el control Data en un nuevo formulario"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NoDataControl"
            Object.ToolTipText     =   "No usar el control Data en un nuevo formulario"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DBGrid"
            Object.ToolTipText     =   "Usar el control DBGrid en un nuevo formulario"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "BeginTrans"
            Object.ToolTipText     =   "Iniciar una transacción"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Rollback"
            Object.ToolTipText     =   "Deshacer la transacción actual"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Commit"
            Object.ToolTipText     =   "Ejecutar la transacción actual"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10005
      TabIndex        =   2
      Top             =   420
      Width           =   10005
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   -15
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FilterIndex     =   1144
      FontSize        =   1.17491e-38
   End
   Begin ComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14579
            Text            =   "Listo"
            TextSave        =   "Listo"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   495
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":014F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":0469
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":0783
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":0A9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":0DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":10D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":13EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":1705
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":1A1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VDMDI.frx":1D39
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Archivo"
      HelpContextID   =   2096095
      Begin VB.Menu mnuDBOpen 
         Caption         =   "&Abrir base de datos..."
         HelpContextID   =   2016062
         Begin VB.Menu mnuDBOMDB 
            Caption         =   "&Microsoft Access..."
         End
         Begin VB.Menu mnuDBOdBASE 
            Caption         =   "&Dbase"
            Begin VB.Menu mnuDBOdBASE5 
               Caption         =   "&5.0..."
            End
            Begin VB.Menu mnuDBOdBASE4 
               Caption         =   "I&V..."
            End
            Begin VB.Menu mnuDBOdBASE3 
               Caption         =   "&III..."
            End
         End
         Begin VB.Menu mnuDBOFoxPro 
            Caption         =   "&FoxPro"
            Begin VB.Menu mnuDBOFox30 
               Caption         =   "&3.0..."
            End
            Begin VB.Menu mnuDBOFox26 
               Caption         =   "2.&6..."
            End
            Begin VB.Menu mnuDBOFox25 
               Caption         =   "2.&5..."
            End
            Begin VB.Menu mnuDBOFox20 
               Caption         =   "2.&0..."
            End
         End
         Begin VB.Menu mnuDBOParadox 
            Caption         =   "&Paradox"
            Begin VB.Menu mnuDBOParadox5 
               Caption         =   "&5.0..."
            End
            Begin VB.Menu mnuDBOParadox4 
               Caption         =   "&4.X..."
            End
            Begin VB.Menu mnuDBOParadox3 
               Caption         =   "&3.X..."
            End
         End
         Begin VB.Menu mnuDBOBtrieve 
            Caption         =   "&Btrieve..."
         End
         Begin VB.Menu mnuDBOExcel 
            Caption         =   "&Excel..."
         End
         Begin VB.Menu mnuDBOText 
            Caption         =   "Archivos de &texto..."
         End
         Begin VB.Menu mnuDBOODBC 
            Caption         =   "&ODBC..."
            HelpContextID   =   2016138
         End
      End
      Begin VB.Menu mnuDBNew 
         Caption         =   "&Nueva..."
         HelpContextID   =   2016083
         Begin VB.Menu mnuDBNMDB 
            Caption         =   "&Microsoft Access"
            Begin VB.Menu mnuDBNMDB2x 
               Caption         =   "MDB Versión &2.0..."
            End
            Begin VB.Menu mnuDBNMDB70 
               Caption         =   "MDB Versión &7.0..."
            End
         End
         Begin VB.Menu mnuDBNdBASE 
            Caption         =   "&Dbase"
            Begin VB.Menu mnuDBNdBASE5 
               Caption         =   "&5.0..."
            End
            Begin VB.Menu mnuDBNdBASE4 
               Caption         =   "I&V..."
            End
            Begin VB.Menu mnuDBNdBASE3 
               Caption         =   "&III..."
            End
         End
         Begin VB.Menu mnuDBNFoxPro 
            Caption         =   "&FoxPro"
            Begin VB.Menu mnuDBNFox30 
               Caption         =   "&3.0..."
            End
            Begin VB.Menu mnuDBNFox26 
               Caption         =   "2.&6..."
            End
            Begin VB.Menu mnuDBNFox25 
               Caption         =   "2.&5..."
            End
            Begin VB.Menu mnuDBNFox20 
               Caption         =   "2.&0..."
            End
         End
         Begin VB.Menu mnuDBNParadox 
            Caption         =   "&Paradox"
            Begin VB.Menu mnuDBNParadox5 
               Caption         =   "&5.0..."
            End
            Begin VB.Menu mnuDBNParadox4 
               Caption         =   "&4.X..."
            End
            Begin VB.Menu mnuDBNParadox3 
               Caption         =   "&3.X..."
            End
         End
         Begin VB.Menu mnuDBNBtrieve 
            Caption         =   "&Btrieve..."
         End
         Begin VB.Menu mnuDBNODBC 
            Caption         =   "&ODBC..."
         End
         Begin VB.Menu mnuDBNText 
            Caption         =   "Archivos de &texto..."
         End
      End
      Begin VB.Menu mnuDBClose 
         Caption         =   "&Cerrar"
         Enabled         =   0   'False
         HelpContextID   =   2016079
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBImpExp 
         Caption         =   "&Importar/Exportar..."
         Enabled         =   0   'False
         HelpContextID   =   2016092
      End
      Begin VB.Menu mnuDBWorkspace 
         Caption         =   "&Espacio de trabajo..."
         HelpContextID   =   2016080
      End
      Begin VB.Menu mnuDBErrors 
         Caption         =   "Err&ores..."
         HelpContextID   =   2016081
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBCompact 
         Caption         =   "Co&mpactar MDB..."
         HelpContextID   =   2016084
         Begin VB.Menu mnuDBC70MDB 
            Caption         =   "&7.0 MDB..."
            HelpContextID   =   2016084
         End
         Begin VB.Menu mnuDBC20MDB 
            Caption         =   "&2.0 MDB..."
            HelpContextID   =   2016084
         End
      End
      Begin VB.Menu mnuDBRepair 
         Caption         =   "&Reparar MDB..."
         HelpContextID   =   2016085
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&1"
         HelpContextID   =   2016095
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&2"
         HelpContextID   =   2016095
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&3"
         HelpContextID   =   2016095
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&4"
         HelpContextID   =   2016095
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&5"
         HelpContextID   =   2016095
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&6"
         HelpContextID   =   2016095
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&7"
         HelpContextID   =   2016095
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBMRU 
         Caption         =   "&8"
         HelpContextID   =   2016095
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBarMRU 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBExit 
         Caption         =   "&Salir"
         HelpContextID   =   2016095
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utili&dad"
      Enabled         =   0   'False
      HelpContextID   =   2096097
      Begin VB.Menu mnuUQuery 
         Caption         =   "&Generador de consultas..."
         HelpContextID   =   2016115
      End
      Begin VB.Menu mnuUDataFormDesigner 
         Caption         =   "&Diseñador de formularios de datos..."
         HelpContextID   =   2098108
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUReplace 
         Caption         =   "&Sustitución global..."
         HelpContextID   =   2016091
      End
      Begin VB.Menu mnuUBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUAttachments 
         Caption         =   "&Adjuntos..."
         HelpContextID   =   2016086
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUGroupsUsers 
         Caption         =   "&Grupos/Usuarios..."
         HelpContextID   =   2016088
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUSystemDB 
         Caption         =   "&SYSTEM.MD?..."
         HelpContextID   =   2016090
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPref 
         Caption         =   "&Preferencias"
         HelpContextID   =   2093354
         Begin VB.Menu mnuPOpenOnStartup 
            Caption         =   "&Abrir la última base de datos al iniciar"
         End
         Begin VB.Menu mnuPAllowSys 
            Caption         =   "&Incluir las tablas del sistema"
         End
         Begin VB.Menu mnuBar4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPQueryTimeout 
            Caption         =   "&Valor de tiempo de espera de la consulta..."
         End
         Begin VB.Menu mnuPLoginTimeout 
            Caption         =   "V&alor de tiempo de espera de inicio de sesión..."
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Ventana"
      HelpContextID   =   2016100
      WindowList      =   -1  'True
      Begin VB.Menu mnuWTile 
         Caption         =   "&Mosaico"
      End
      Begin VB.Menu mnuWCascade 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "&Organizar iconos"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      HelpContextID   =   2093307
      Begin VB.Menu mnuHSearch 
         Caption         =   "&Buscar..."
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuDBPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDBPUOpen 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuDBPUDesign 
         Caption         =   "&Diseñar..."
      End
      Begin VB.Menu mnuDBPUEdit 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnuDBPURename 
         Caption         =   "&Renombrar"
      End
      Begin VB.Menu mnuDBPUDelete 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuDBPUBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBPUCopyStruct 
         Caption         =   "Copiar estructura..."
      End
      Begin VB.Menu mnuDBPURemoveAll 
         Caption         =   "Eliminar todos los registros"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDBPURefresh 
         Caption         =   "Actualizar lista"
      End
      Begin VB.Menu mnuDBPUBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBPUNewTable 
         Caption         =   "Nue&va tabla"
      End
      Begin VB.Menu mnuDBPUNewQuery 
         Caption         =   "Nueva &consulta"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary
'>>>>>>>>>>>>>>>>>>>>>>>>
'menús
Const MNU_Database = "&Archivo"
Const MNU_DBOpen = "&Abrir base de datos..."
Const MNU_DBOMDB = "&Microsoft Access..."
Const MNU_DBOText = "Archivos de &texto..."
Const MNU_DBOODBC = "&ODBC..."
Const MNU_DBNew = "&Nueva..."
Const MNU_DBNMDB = "&Microsoft Access"
Const MNU_DBNMDB2x = "MDB Versión &2.0..."
Const MNU_DBNMDB70 = "MDB Versión &7.0..."
Const MNU_DBNODBC = "&ODBC..."
Const MNU_DBNText = "Archivos de &texto..."
Const MNU_DBClose = "&Cerrar"
Const MNU_DBImpExp = "&Importar/Exportar..."
Const MNU_DBWorkspace = "&Espacio de trabajo..."
Const MNU_DBErrors = "Err&ores..."
Const MNU_DBCompact = "Co&mpactar MDB..."
Const MNU_DBRepair = "&Reparar MDB..."
Const MNU_DBExit = "&Salir"
Const MNU_Util = "Utili&dad"
Const MNU_UQuery = "&Generador de consultas..."
Const MNU_UDataFormDesigner = "&Diseñador de formularios de datos..."
Const MNU_UReplace = "&Sustitución global..."
Const MNU_UAttachments = "&Adjuntos..."
Const MNU_UGroupsUsers = "&Grupos/Usuarios..."
Const MNU_USystemDB = "&SYSTEM.MD?..."
Const MNU_Pref = "&Preferencias"
Const MNU_POpenOnStartup = "&Abrir la última base de datos al iniciar"
Const MNU_PAllowSys = "&Incluir las tablas del sistema"
Const MNU_PQueryTimeout = "&Valor de tiempo de espera de la consulta..."
Const MNU_PLoginTimeout = "V&alor de tiempo de espera de inicio de sesión..."
Const MNU_Window = "&Ventana"
Const MNU_WTile = "&Mosaico"
Const MNU_WCascade = "&Cascada"
Const MNU_WArrange = "&Organizar iconos"
Const MNU_Help = "Ay&uda"
Const MNU_HSearch = "Í&ndice de referencia de búsqueda..."
Const MNU_HAbout = "&Acerca de..."
Const MNU_DBPUOpen = "&Abrir"
Const MNU_DBPUDesign = "&Diseñar..."
Const MNU_DBPUEdit = "&Editar"
Const MNU_DBPURename = "&Renombrar"
Const MNU_DBPUDelete = "&Eliminar"
Const MNU_DBPUCopyStruct = "Copiar estructura..."
Const MNU_DBPURemoveAll = "Eliminar todos los registros"
Const MNU_DBPURefresh = "Actualizar lista"
Const MNU_DBPUNewTable = "Nue&va tabla"
Const MNU_DBPUNewQuery = "Nueva &consulta"

'información sobre herramientas
Const TOOLTIP1 = "Recordset de tipo Table"
Const TOOLTIP2 = "Recordset de tipo Dynaset"
Const TOOLTIP3 = "Recordset de tipo Snapshot"
Const TOOLTIP4 = "Recordset de tipo Passthrough"
Const TOOLTIP5 = "Usar el control Data en un nuevo formulario"
Const TOOLTIP6 = "No usar el control Data en un nuevo formulario"
Const TOOLTIP7 = "Usar el control DBGrid en un nuevo formulario"
Const TOOLTIP8 = "Iniciar una transacción"
Const TOOLTIP9 = "Deshacer la transacción actual"
Const TOOLTIP10 = "Ejecutar la transacción actual"

'cadenas variadas
Const MSG3 = "Presione cualquier tecla para cerrar el cuadro de diálogo Acerca de"
Const MSG4 = "Escriba los nuevos parámetros de la base de datos"
Const MSG5 = "Introduzca el nombre del controlador desde el archivo ODBCINST.INI:"
Const MSG6 = "Nombre del controlador"
Const MSG7 = "Debe cerrar primero"
Const MSG8 = "NOTA: Es recomendable usar tablas adjuntas"
Const MSG9 = "MDB de Microsoft Access (*.mdb)|*.mdb|Todos los archivos (*.*)|*.*"
Const MSG10 = "Abrir la base de datos de Microsoft Access para reparar"
Const MSG11 = "Reparando "
Const MSG12 = "¿Desea abrir la base de datos reparada?"
Const MSG13 = "Bases de datos del sistema|SYSTEM.MD?"
Const MSG14 = "Seleccionar SYSTEM.MD? (Archivo de seguridad de Microsoft Access)"
Const MSG15 = " Usuario: "
Const MSG16 = "Es necesario cerrar la base de datos actual por un error"
Const MSG17 = "No se encontraron usuarios, intente 'Utilidad/System MD?'"
Const MSG18 = "Tiempo de inicio de sesión (en segundos):"
Const MSG19 = "No hay ninguna base de datos abierta"
Const MSG20 = "Tiempo de inicio de consulta (en segundos):"
Const MSG21 = "¿Desea eliminar la tabla?"
Const MSG22 = "¿Desea eliminar el QueryDef?"
Const MSG23 = "¿Desea eliminar el campo?"
Const MSG24 = "¿Desea eliminar el índice?"
Const MSG25 = "¿Desea eliminar todos los registros en la tabla?"
Const MSG26 = "Filas eliminadas: "
Const MSG27 = "No se encontró SYSTEM.MD?, ¿Desea agregar uno a la configuración de VB?"
Const MSG28 = "Este controlador no es compatible con transacciones"
Const MSG29 = "Se perderán todos los cambios, ¿desea continuar?"
Const MSG30 = " La propiedad es de sólo lectura"
Const MSG31 = "Esta función necesita un proyecto activo"
'>>>>>>>>>>>>>>>>>>>>>>>>

Dim mHwnd As Long

Private Sub mnuDBC70MDB_Click()
  CompactDB dbVersion30
End Sub

Private Sub mnuDBNMDB2x_Click()
  NewMDB dbVersion20
End Sub

Private Sub mnuDBNMDB70_Click()
  NewMDB dbVersion30
End Sub

Private Sub mnuDBOExcel_Click()
  'podemos utilizar Excel 5.0 con todos los archivos de Excel
  'puesto que el ISAM detectará la versión al
  'abrir el archivo
  gsDataType = gsEXCEL50
  OpenLocalDB False
End Sub

Private Sub mnuDBPUDesign_Click()
  On Error Resume Next

  If gnodDBNode2 Is Nothing Then Exit Sub
  If gnodDBNode2.Tag = TABLE_STR Then
    gbAddTableFlag = False
    Screen.MousePointer = vbHourglass
    frmTblStruct.Show vbModal
  ElseIf gnodDBNode2.Tag = QUERY_STR Then
    Set frmDatabase.tvDatabase.SelectedItem = gnodDBNode2
    frmSQL.txtSQLStatement.Text = gdbCurrentDB.QueryDefs(gnodDBNode2.Text).SQL
  End If

End Sub

Sub mnuDBPUEdit_Click()
  On Error GoTo mnuDBPUEdit_ClickErr
  
  Dim prpObj As Property
  Dim vTmp As Variant
  Dim vNew As Variant
  Dim frmProp As New frmProperty
  
  If gnodDBNode2.Parent.Parent Is Nothing Then
    'tiene que ser una propiedad de la base de datos
    Set prpObj = gdbCurrentDB.Properties(VBA.Left(gnodDBNode2.Text, InStr(gnodDBNode2.Text, "=") - 1))
  Else
    Select Case gnodDBNode2.Parent.Parent.Tag
      Case TABLE_STR
        Set prpObj = gdbCurrentDB.TableDefs(gnodDBNode2.Parent.Parent.Text).Properties(VBA.Left(gnodDBNode2.Text, InStr(gnodDBNode2.Text, "=") - 1))
      Case QUERY_STR
        Set prpObj = gdbCurrentDB.QueryDefs(gnodDBNode2.Parent.Parent.Text).Properties(VBA.Left(gnodDBNode2.Text, InStr(gnodDBNode2.Text, "=") - 1))
      Case FIELDS_STR
        Set prpObj = gdbCurrentDB.TableDefs(gnodDBNode2.Parent.Parent.Parent.Text).Fields(gnodDBNode2.Parent.Text).Properties(VBA.Left(gnodDBNode2.Text, InStr(gnodDBNode2.Text, "=") - 1))
      Case INDEXES_STR
        Set prpObj = gdbCurrentDB.TableDefs(gnodDBNode2.Parent.Parent.Parent.Text).Indexes(gnodDBNode2.Parent.Text).Properties(VBA.Left(gnodDBNode2.Text, InStr(gnodDBNode2.Text, "=") - 1))
      Case Else
        Exit Sub
    End Select
  End If
  
  'almacena el valor
  vTmp = prpObj.Value
  
  On Error Resume Next
  'intenta establecerlo a su valor actual
  'para ver si es de sólo lectura
  prpObj.Value = vTmp
  If Err Then
    'es de sólo lectura, así que sale
    Err.Clear
    MsgBox "'" & prpObj.Name & "': " & MSG30, vbExclamation
    Exit Sub
  End If
  On Error GoTo mnuDBPUEdit_ClickErr
  
  With frmProp
    Set .PropObject = prpObj
    .Show vbModal
    If .OK Then
      gnodDBNode2.Text = prpObj.Name & "=" & prpObj.Value
      'Comprueba si es una propiedad Name
      If prpObj.Name = "Name" Then
        gnodDBNode2.Parent.Text = prpObj.Value
      End If
    End If
    Unload frmProp
  End With
  Set frmProp = Nothing
  
  Exit Sub
mnuDBPUEdit_ClickErr:
  ShowError
End Sub

Private Sub mnuDBPUNewQuery_Click()
  'carga el formulario de Consultas para generar la consulta
  Unload frmQuery      'despeja la pantalla
  frmQuery.Show
End Sub

Private Sub mnuDBPUNewTable_Click()
  gbAddTableFlag = True
  Screen.MousePointer = vbHourglass
  frmTblStruct.Show vbModal
End Sub

Sub mnuDBPUOpen_Click()
  On Error Resume Next
  
  gbFromSQL = False
  Screen.MousePointer = vbHourglass
  
  If gnodDBNode.Tag = TABLE_STR Then
    OpenTable StripConnect(gnodDBNode2.Text)
  ElseIf gnodDBNode.Tag = QUERY_STR Then
    OpenQuery gnodDBNode2.Text, False
  End If
  
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False

End Sub

Private Sub mnuHAbout_Click()
  MsgBar MSG3, False
  frmAboutBox.Show vbModal
  MsgBar vbNullString, False
End Sub

Private Sub mnuDBC20MDB_Click()
  CompactDB dbVersion20
End Sub

Private Sub mnuDBClose_Click()
  CloseCurrentDB
End Sub

Private Sub mnuDBErrors_Click()
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  RefreshErrors
  Screen.MousePointer = vbDefault
  If Err Then ShowError
End Sub

Private Sub mnuDBExit_Click()
  Unload Me
End Sub

Private Sub mnuDBNBtrieve_Click()
   gsDataType = gsBTRIEVE
   NewLocalISAM
End Sub

Private Sub mnuDBNDbase3_Click()
   gsDataType = gsDBASEIII
   NewLocalISAM
End Sub

Private Sub mnuDBNDbase4_Click()
   gsDataType = gsDBASEIV
   NewLocalISAM
End Sub

Private Sub mnuDBNDbase5_Click()
   gsDataType = gsDBASE5
   NewLocalISAM
End Sub

Private Sub mnuDBNFox20_Click()
   gsDataType = gsFOXPRO20
   NewLocalISAM
End Sub

Private Sub mnuDBNFox25_Click()
   gsDataType = gsFOXPRO25
   NewLocalISAM
End Sub

Private Sub mnuDBNFox26_Click()
   gsDataType = gsFOXPRO26
   NewLocalISAM
End Sub

Private Sub mnuDBNFox30_Click()
   gsDataType = gsFOXPRO30
   NewLocalISAM
End Sub

Private Sub mnuDBNODBC_Click()
  On Error GoTo DBNErr

  Dim sDriverName As String

  MsgBar MSG4, False

  'el controlador tiene que tener una entrada válida en ODBCINST.INI
  sDriverName = InputBox(MSG5, MSG6, gsDEFAULT_DRIVER)

  If Len(sDriverName) = 0 Then Exit Sub 'han cancelado
  
  DBEngine.RegisterDatabase vbNullString, sDriverName, False, vbNullString

  SendKeys "%AAO"   'fuerza a mostrarse el cuadro de diálogo Abrir base de datos

  MsgBar vbNullString, False
  Exit Sub

DBNErr:
  ShowError
End Sub

Private Sub mnuDBNParadox3_Click()
  gsDataType = gsPARADOX3X
  NewLocalISAM
End Sub

Private Sub mnuDBNParadox4_Click()
  gsDataType = gsPARADOX4X
  NewLocalISAM
End Sub

Private Sub mnuDBNParadox5_Click()
  gsDataType = gsPARADOX5X
  NewLocalISAM
End Sub

Private Sub mnuDBNText_Click()
  gsDataType = gsTEXTFILES
  NewLocalISAM
End Sub

Private Sub mnuDBOMDB_Click()
  gsDataType = gsMSACCESS
  OpenLocalDB False
End Sub

Private Sub mnuDBOBtrieve_Click()
  gsDataType = gsBTRIEVE
  OpenLocalDB False
End Sub

Private Sub mnuDBODbase3_Click()
   gsDataType = gsDBASEIII
   OpenLocalDB False
End Sub

Private Sub mnuDBODbase4_Click()
   gsDataType = gsDBASEIV
   OpenLocalDB False
End Sub

Private Sub mnuDBODbase5_Click()
   gsDataType = gsDBASE5
   OpenLocalDB False
End Sub

Private Sub mnuDBOFox20_Click()
   gsDataType = gsFOXPRO20
   OpenLocalDB False
End Sub

Private Sub mnuDBOFox25_Click()
   gsDataType = gsFOXPRO25
   OpenLocalDB False
End Sub

Private Sub mnuDBOFox26_Click()
   gsDataType = gsFOXPRO26
   OpenLocalDB False
End Sub

Private Sub mnuDBOFox30_Click()
   gsDataType = gsFOXPRO30
   OpenLocalDB False
End Sub

Private Sub mnuDBOODBC_Click()
  Dim frm As New frmODBCLogon

  frm.Show vbModal
  
  If frm.DBOpened Then
    ShowDBTools
    RefreshTables Nothing
    MsgBar MSG8, False
  End If
  
  Unload frm
  Set frm = Nothing
  
End Sub

Private Sub mnuDBOParadox3_Click()
   gsDataType = gsPARADOX3X
   OpenLocalDB False
End Sub

Private Sub mnuDBOParadox4_Click()
   gsDataType = gsPARADOX4X
   OpenLocalDB False
End Sub

Private Sub mnuDBOParadox5_Click()
   gsDataType = gsPARADOX5X
   OpenLocalDB False
End Sub

Private Sub mnuDBOText_Click()
   gsDataType = gsTEXTFILES
   OpenLocalDB False
End Sub

Private Sub mnuDBRepair_Click()
  On Error GoTo RepairAccErr

  Dim sNewName As String

  'obtiene el nombre del archivo a reparar
  With dlgCMD1
    .Filter = MSG9
    .DialogTitle = MSG10
    .FilterIndex = 1
    .Flags = FileOpenConstants.cdlOFNHideReadOnly
    .ShowOpen
  End With
  If Len(dlgCMD1.FileName) > 0 Then
    sNewName = dlgCMD1.FileName
  Else
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  MsgBar MSG11 & sNewName, True
  DBEngine.RepairDatabase sNewName
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False

  If MsgBox(MSG12, vbYesNo + vbQuestion) = vbYes Then
    If gbDBOpenFlag Then
      Call mnuDBClose_Click
    End If
    gsDataType = gsMSACCESS
    gsDBName = sNewName
    OpenLocalDB True
  End If

  If gbDBOpenFlag Then
    ShowDBTools
    RefreshTables Nothing
  End If

  Exit Sub

RepairAccErr:
  If Err <> 32755 Then
    ShowError
  End If
End Sub

Private Sub mnuHSearch_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, HelpConstants.cdlHelpPartialKey, 0)
  If Err Then
    ShowError
  End If
End Sub

Private Sub mnuUSystemDB_Click()
  On Error Resume Next
  
  Dim sTmp As String
  Dim x As Integer
  
  With dlgCMD1
    .Filter = MSG13
    .DialogTitle = MSG14
    .FilterIndex = 1
    .FileName = "SYSTEM.MDW"
    .CancelError = True
    .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
  End With
  On Error Resume Next
  dlgCMD1.ShowOpen
  If Err = 32755 Then         'el usuario ha cancelado
    Exit Sub
  Else
    sTmp = dlgCMD1.FileName  'tiene que ser un nombre de archivo bueno
    SaveSetting APP_CATEGORY & "\VisData", "Engines", "SystemDB", sTmp
    SaveSetting APP_CATEGORY, App.Title, "LoadSystemDB", "Yes"
  End If

End Sub

Private Sub mnuDBWorkspace_Click()
  On Error GoTo WSErr
    
  Dim sDBName As String
  Dim sConnect As String
  Dim sUser As String
    
  If gbDBOpenFlag Then
    'guarda los valores anteriores
    sDBName = gdbCurrentDB.Name
    sConnect = gdbCurrentDB.Connect
    sUser = gwsMainWS.UserName
  End If
  
  frmLogin.Show vbModal
  stsStatusBar.Panels(2).Text = MSG15 & gwsMainWS.UserName & " "
  
  'vuelve a abrir la base de datos si hay cambio de usuario
  If UCase(sUser) <> UCase(gwsMainWS.UserName) And gbDBOpenFlag Then
    'tiene que cerrar los objetos no válidos después de volver a abrir la base de datos
    CloseAllRecordsets
    Set gdbCurrentDB = gwsMainWS.OpenDatabase(sDBName, False, gnReadOnly, sConnect)
  End If
  
  Exit Sub
  
WSErr:
  ShowError
  If gbDBOpenFlag Then
    MsgBox MSG16, 48
  End If
  Call mnuDBClose_Click
End Sub

Private Sub mnuUAttachments_Click()
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  frmAttachments.Show
  Screen.MousePointer = vbDefault
  If Err Then ShowError
End Sub

Private Sub mnuUGroupsUsers_Click()
  On Error Resume Next
    
  If gwsMainWS.Users.Count = 0 Then
    Beep
    MsgBox MSG17, 48
    Exit Sub
  End If
    
  Screen.MousePointer = vbHourglass
  frmGroupsUsers.Show
  Screen.MousePointer = vbDefault
  If Err Then ShowError
End Sub

Private Sub mnuPAllowSys_Click()
  On Error Resume Next

  mnuPAllowSys.Checked = Not mnuPAllowSys.Checked
  If Not gbDBOpenFlag Then Exit Sub
  
  RefreshTables Nothing
  
End Sub

Private Sub mnuPLoginTimeout_Click()
  On Error GoTo LTErr

  Dim sNewValue As String

  sNewValue = InputBox(MSG18, , CStr(glLoginTimeout))
  If Len(sNewValue) = 0 Then Exit Sub

  'intenta establecer el nuevo valor
  If Val(sNewValue) >= 0 Then
    glLoginTimeout = Val(sNewValue)
    DBEngine.LoginTimeout = glLoginTimeout
  End If

  Exit Sub

LTErr:
  ShowError
End Sub

Private Sub mnuPOpenOnStartup_Click()
  mnuPOpenOnStartup.Checked = Not mnuPOpenOnStartup.Checked
End Sub

Private Sub mnuPQueryTimeout_Click()
  On Error GoTo QTErr

  Dim sNewValue As String

  If Not gbDBOpenFlag Then MsgBox MSG19, 48: Exit Sub
  
  sNewValue = InputBox(MSG20, , CStr(gdbCurrentDB.QueryTimeout))
  If Len(sNewValue) = 0 Then Exit Sub

  'intenta establecer el nuevo valor
  gdbCurrentDB.QueryTimeout = Val(sNewValue)
  glQueryTimeout = Val(sNewValue)

  Exit Sub

QTErr:
  ShowError
  'restablece el control formulario después del error
  glQueryTimeout = gdbCurrentDB.QueryTimeout
End Sub

Private Sub mnuUDataFormDesigner_Click()
  On Error Resume Next
  
  'asegura que hay un proyecto cargado
  If gVDClass.VBInstance.ActiveVBProject Is Nothing Then
    MsgBox MSG31, vbInformation
    Exit Sub
  End If
  
  frmDFD.Show vbModal
  If Err Then ShowError
End Sub

Private Sub mnuUQuery_Click()
  frmQuery.Show
  frmQuery.WindowState = 0
End Sub

Private Sub mnuDBPUCopyStruct_Click()
  On Error Resume Next
  frmCopyStruct.Show vbModal
  If Err Then ShowError
End Sub

Private Sub mnuDBPUDelete_Click()
  On Error GoTo TblDelErr

  Dim sName As String
  If gnodDBNode2 Is Nothing Then Exit Sub
  
  Select Case gnodDBNode2.Tag
    Case TABLE_STR
      sName = StripConnect(gnodDBNode2.Text)
      If MsgBox(MSG21, vbYesNo + vbQuestion) = vbYes Then
        gdbCurrentDB.TableDefs.Delete sName
        frmDatabase.tvDatabase.Nodes.Remove gnodDBNode2.Index
      End If
    Case QUERY_STR
      sName = gnodDBNode2.Text
      If MsgBox(MSG22, vbYesNo + vbQuestion) = vbYes Then
        gdbCurrentDB.QueryDefs.Delete sName
        frmDatabase.tvDatabase.Nodes.Remove gnodDBNode2.Index
      End If
    Case FIELD_STR
      sName = gnodDBNode2.Text
      If MsgBox(MSG23, vbYesNo + vbQuestion) = vbYes Then
        gdbCurrentDB.TableDefs(gnodDBNode2.Parent.Parent.Text).Fields.Delete sName
        frmDatabase.tvDatabase.Nodes.Remove gnodDBNode2.Index
      End If
    Case INDEX_STR
      sName = gnodDBNode2.Text
      If MsgBox(MSG24, vbYesNo + vbQuestion) = vbYes Then
        gdbCurrentDB.TableDefs(gnodDBNode2.Parent.Parent.Text).Indexes.Delete sName
        frmDatabase.tvDatabase.Nodes.Remove gnodDBNode2.Index
      End If
  End Select

  Exit Sub

TblDelErr:
  ShowError
End Sub

Private Sub mnuDBPURefresh_Click()
  gdbCurrentDB.TableDefs.Refresh
  RefreshTables Nothing
End Sub

Private Sub mnuDBPURename_Click()
  On Error GoTo mnuDBPURename_ClickErr
  
  If Not gnodDBNode2 Is Nothing Then
    'lo establece al nuevo nodo para editar
    Set frmDatabase.tvDatabase.SelectedItem = gnodDBNode2
    frmDatabase.tvDatabase.StartLabelEdit
  End If

  Exit Sub

mnuDBPURename_ClickErr:
  ShowError
End Sub

Private Sub mnuDBPURemoveAll_Click()
  On Error GoTo RemoveAllErr

  Dim sTBLName As String
  
  sTBLName = StripConnect(gnodDBNode.Text)

  If MsgBox(MSG25 & " '" & sTBLName & "'", vbYesNo + vbQuestion) = vbYes Then
    'elimina todas las filas con una instrucción Sql
    If gsDataType = gsSQLDB Then
      gdbCurrentDB.Execute ("delete from " & sTBLName), dbSQLPassThrough
    Else
      gdbCurrentDB.Execute ("delete from " & sTBLName)
    End If
    If gdbCurrentDB.RecordsAffected > 0 Then
      MsgBox MSG26 & gdbCurrentDB.RecordsAffected, 48
      If gbTransPending Then gbDBChanged = True
    End If

  End If

  Exit Sub

RemoveAllErr:
  If Err = gnEOF_ERR Then Resume Next
  ShowError
End Sub

Private Sub mnuDBImpExp_Click()
  On Error Resume Next
  frmImpExp.Show vbModal
  If Err Then ShowError
End Sub

Private Sub mnuUReplace_Click()
  On Error GoTo ReplaceErr

  frmReplace.Show vbModal

  Exit Sub

ReplaceErr:
  ShowError
End Sub

Private Sub mnuWArrange_Click()
  Me.Arrange 3
End Sub

Private Sub mnuWCascade_Click()
  Me.Arrange 0
End Sub

Private Sub mnuWTile_Click()
  Me.Arrange 2
End Sub

Private Sub MDIForm_Load()
  On Error GoTo MDILErr

  Dim x As Integer

  'carga cadenas desde las constantes de
  'menús
  mnuDatabase.Caption = MNU_Database
  mnuDBOpen.Caption = MNU_DBOpen
  mnuDBOMDB.Caption = MNU_DBOMDB
  mnuDBOText.Caption = MNU_DBOText
  mnuDBOODBC.Caption = MNU_DBOODBC
  mnuDBNew.Caption = MNU_DBNew
  mnuDBNMDB.Caption = MNU_DBNMDB
  mnuDBNMDB2x.Caption = MNU_DBNMDB2x
  mnuDBNMDB70.Caption = MNU_DBNMDB70
  mnuDBNODBC.Caption = MNU_DBNODBC
  mnuDBNText.Caption = MNU_DBNText
  mnuDBClose.Caption = MNU_DBClose
  mnuDBImpExp.Caption = MNU_DBImpExp
  mnuDBWorkspace.Caption = MNU_DBWorkspace
  mnuDBErrors.Caption = MNU_DBErrors
  mnuDBCompact.Caption = MNU_DBCompact
  mnuDBRepair.Caption = MNU_DBRepair
  mnuDBExit.Caption = MNU_DBExit
  mnuUtil.Caption = MNU_Util
  mnuUQuery.Caption = MNU_UQuery
  mnuUDataFormDesigner.Caption = MNU_UDataFormDesigner
  mnuUReplace.Caption = MNU_UReplace
  mnuUAttachments.Caption = MNU_UAttachments
  mnuUGroupsUsers.Caption = MNU_UGroupsUsers
  mnuUSystemDB.Caption = MNU_USystemDB
  mnuPref.Caption = MNU_Pref
  mnuPOpenOnStartup.Caption = MNU_POpenOnStartup
  mnuPAllowSys.Caption = MNU_PAllowSys
  mnuPQueryTimeout.Caption = MNU_PQueryTimeout
  mnuPLoginTimeout.Caption = MNU_PLoginTimeout
  mnuWindow.Caption = MNU_Window
  mnuWTile.Caption = MNU_WTile
  mnuWCascade.Caption = MNU_WCascade
  mnuWArrange.Caption = MNU_WArrange
  mnuHelp.Caption = MNU_Help
  mnuHSearch.Caption = MNU_HSearch
  mnuHAbout.Caption = MNU_HAbout
  mnuDBPUOpen.Caption = MNU_DBPUOpen
  mnuDBPUDesign.Caption = MNU_DBPUDesign
  mnuDBPUEdit.Caption = MNU_DBPUEdit
  mnuDBPURename.Caption = MNU_DBPURename
  mnuDBPUDelete.Caption = MNU_DBPUDelete
  mnuDBPUCopyStruct.Caption = MNU_DBPUCopyStruct
  mnuDBPURemoveAll.Caption = MNU_DBPURemoveAll
  mnuDBPURefresh.Caption = MNU_DBPURefresh
  mnuDBPUNewTable.Caption = MNU_DBPUNewTable
  mnuDBPUNewQuery.Caption = MNU_DBPUNewQuery
  
  'información de herramientas
  tlbToolBar.Buttons(1).ToolTipText = TOOLTIP1
  tlbToolBar.Buttons(2).ToolTipText = TOOLTIP2
  tlbToolBar.Buttons(3).ToolTipText = TOOLTIP3
  tlbToolBar.Buttons(4).ToolTipText = TOOLTIP4
  tlbToolBar.Buttons(6).ToolTipText = TOOLTIP5
  tlbToolBar.Buttons(7).ToolTipText = TOOLTIP6
  tlbToolBar.Buttons(8).ToolTipText = TOOLTIP7
  tlbToolBar.Buttons(10).ToolTipText = TOOLTIP8
  tlbToolBar.Buttons(11).ToolTipText = TOOLTIP9
  tlbToolBar.Buttons(12).ToolTipText = TOOLTIP10

  gnMULocking = True   'bloqueo pesimista predeterminado
  App.HelpFile = App.Path & "\HELP\VB5.HLP"
 
  'hay que deshabilitar los elementos de menú de Btrieve en 32 bits
  mnuDBOBtrieve.Visible = False
  mnuDBNBtrieve.Visible = False

  'obtiene las coordenadas del formulario
  x = Val(GetINIString("WindowState", "2"))
  If x <> 1 Then
    frmMDI.WindowState = x
  Else
    frmMDI.WindowState = 0
  End If
  If frmMDI.WindowState = 0 Then
    frmMDI.Left = Val(GetINIString("WindowLeft", "0"))
    frmMDI.Top = Val(GetINIString("WindowTop", "0"))
    frmMDI.Width = Val(GetINIString("WindowWidth", "9135"))
    frmMDI.Height = Val(GetINIString("WindowHeight", "6900"))
  End If
  
  'comprueba si el usuario respondió no al agregar system.mda
  If Len(GetINIString("LoadSystemDB", vbNullString)) = 0 Then
    'es la primera vez, luego pide que se agregue si no está presente
    If MsgBox("¿Desea agregar SYSTEM.MD? (Archivo de seguridad de Microsoft Access) al archivo INI?", vbYesNo + vbQuestion) = vbYes Then
      mnuUSystemDB_Click
    Else
      'almacena la información y no seguimos preguntando
      SaveSetting APP_CATEGORY, App.Title, "LoadSystemDB", "No"
    End If
  End If
  
  On Error GoTo MDILErr
  
  'configura DBEngine
  DBEngine.IniPath = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\" & APP_CATEGORY & "\" & App.Title
  DBEngine.DefaultUser = "admin"
  DBEngine.DefaultPassword = vbNullString

  'inicio de sesión en Jet
  On Error Resume Next
  Set gwsMainWS = DBEngine.CreateWorkspace("MainWS", "admin", vbNullString)
  If Err = 3029 Then
    frmLogin.Show vbModal
  ElseIf Err = 3044 Then  'ruta no válida, system.mda es falso
    If MsgBox(MSG27, vbYesNo + vbQuestion) = vbYes Then
      mnuUSystemDB_Click
    Else
      'almacena la información y no seguimos preguntando
      SaveSetting APP_CATEGORY, App.Title, "LoadSystemDB", "No"
      SaveSetting "VisData", "Options", "SystemDB", vbNullString
    End If
  ElseIf Err <> 0 Then
    ShowError
  End If
  stsStatusBar.Panels(2).Text = MSG15 & gwsMainWS.UserName & " "
  
  On Error GoTo MDILErr
  
  'agrega el entorno de trabajo a la colección
  Workspaces.Append gwsMainWS
  Me.Show
  LoadINISettings
  
  'intenta abrir la última base de datos si dicha opción
  'ha sido establecida en el menú Preferencias
  If frmMDI.mnuPOpenOnStartup.Checked And Len(gsDBName) > 0 Then
    If gsDataType = gsSQLDB Then
'      'En una base de datos ODBC, tenemos que
'      'utilizar Sendkeys para abrir el cuadro de diálogo ODBC
'      SendKeys "%AAO{Enter}"
      mnuDBOODBC_Click
    Else
      OpenLocalDB True
    End If
  Else
    HideDBTools
  End If

  Exit Sub

MDILErr:
  ShowError
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  ShutDownVisData
  
  If mHwnd <> 0 Then
    'necesario cuando VisData ha sido cargado desde el menú Complementos de VB
    mHwnd = SetWindowLong(Me.hwnd, -8, GetDesktopWindow())
  End If
  
End Sub

Private Sub mnuDBMRU_Click(Index As Integer)
  On Error GoTo MRUErr

  gsDBName = Mid(mnuDBMRU(Index).Caption, 4, Len(mnuDBMRU(Index).Caption))
  gsDataType = mnuDBMRU(Index).Tag
  If UCase(Left(gsDataType, 5)) <> gsSQLDB Then
    OpenLocalDB True
  Else
    'tiene que ser una base de datos ODBC, así que tenemos que cargar frmOpenDB
    'obtendremos los componentes de la conexión
    GetODBCConnectParts gsDataType
    'llama a la subrutina que carga el formulario
    mnuDBOODBC_Click
  End If

  Exit Sub

MRUErr:
  ShowError
End Sub

Private Sub tlbToolBar_ButtonClick(ByVal BUTTON As BUTTON)
  On Error GoTo tlbToolBar_ButtonClickErr
  
  Select Case BUTTON.Key
    Case "DataControl"
      gnFormType = gnFORM_DATACTL
    Case "NoDataControl"
      gnFormType = gnFORM_NODATACTL
    Case "DBGrid"
      gnFormType = gnFORM_DATAGRID
    
    Case "Table"
      gnRSType = gnRS_TABLE
    Case "Dynaset"
      gnRSType = gnRS_DYNASET
    Case "Snapshot"
      gnRSType = gnRS_SNAPSHOT
    Case "PassThrough"
      gnRSType = gnRS_PASSTHRU
    
    Case "BeginTrans"
      If gdbCurrentDB.Transactions = False Then
        Beep
        MsgBox MSG28
        Exit Sub
      End If
      gwsMainWS.BeginTrans
      gbDBChanged = False
      gbTransPending = True
      tlbToolBar.Buttons("BeginTrans").Enabled = False
      tlbToolBar.Buttons("Commit").Enabled = True
      tlbToolBar.Buttons("Rollback").Enabled = True
    Case "Deshacer"
      If MsgBox(MSG29, vbYesNo + vbQuestion) = vbYes Then
        gwsMainWS.Rollback
        gbDBChanged = False
        gbTransPending = False
        tlbToolBar.Buttons("BeginTrans").Enabled = True
        tlbToolBar.Buttons("Commit").Enabled = False
        tlbToolBar.Buttons("Rollback").Enabled = False
      End If
    Case "Ejecutar"
      gwsMainWS.CommitTrans
      gbDBChanged = False
      gbTransPending = False
      tlbToolBar.Buttons("BeginTrans").Enabled = True
      tlbToolBar.Buttons("Commit").Enabled = False
      tlbToolBar.Buttons("Rollback").Enabled = False
  End Select
  
  Exit Sub
tlbToolBar_ButtonClickErr:
  ShowError
End Sub

Public Sub SetWindowParent()
  'necesario cuando VisData se carga desde el menú Complementos de VB
  mHwnd = SetWindowLong(Me.hwnd, -8, gVDClass.VBInstance.MainWindow.hwnd)
End Sub

