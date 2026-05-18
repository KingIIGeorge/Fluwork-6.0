Attribute VB_Name = "modVisData"
'------------------------------------------------------------
' VISDATA.BAS
' funciones de soporte de la aplicación de ejemplo Visual Data
'
' Información general: Esta aplicación ha sido diseñada para demostrar
'   y ejercitar toda la funcionalidad disponible en
'   DAO (Objetos para acceso a datos) de Visual Basic 5.0.
'
'------------------------------------------------------------

Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const MSG1 = "Ejecutar primero Commit o Rollback."
Const MSG2 = "Cerrando los Recordsets"
Const MSG3 = "La tabla ya existe, ¿desea eliminarla?"
Const MSG4 = "Introduzca el nuevo nombre de la tabla:"
Const MSG5 = "Listo"
Const MSG6 = ", espere por favor..."
Const MSG7 = "Actualizando la lista de tablas"
Const MSG8 = "Número: "
Const MSG9 = "¿Desea mostrar la colección de errores de acceso a datos?"
Const MSG10 = "Imposible usar un objeto Table en una tabla adjunta, ¿dese usar un Dynaset?"
Const MSG11 = "Abriendo la tabla adjunta como un Dynaset"
Const MSG12 = "Abriendo la tabla adjunta como un Snapshot"
Const MSG13 = "Abriendo la tabla completa"
Const MSG14 = "Abriendo un único Dynaset de tipo Table"
Const MSG15 = "Abriendo un único Snapshot de tipo Table"
Const MSG16 = "Abriendo un Snapshot de tipo PassThru"
Const MSG17 = "¿Es una consulta SQLPassThrough?"
Const MSG18 = "Escriba el valor de la propiedad Connect:"
Const MSG19 = "Imposible abrir un objeto Table de un QueryDef, ¿desea usar un Dynaset?"
Const MSG20 = "Abriendo una consulta de tipo Snapshot"
Const MSG21 = "Abriendo una consulta de tipo Dynaset"
Const MSG22 = "Instrucción SQL"
Const MSG23 = "Ejecutar "
Const MSG24 = " consulta?"
Const MSG25 = "Ejecutando consulta"
Const MSG26 = "  [No actualizable]"
Const MSG27 = "La tabla ya existe, ¿desea eliminarla?"
Const MSG28 = "La consulta ya existe, ¿desea eliminarla?"
Const MSG29 = "Escriba un valor para el parámetro:"
Const MSG30 = "No hay errores actuales de acceso a datos"
Const MSG31 = "Imposible mostrar los errores en este momento"
Const MSG32 = "Los datos han cambiado, ¿desea ejecutar?"
Const MSG33 = "¿Desea deshacer todos los cambios?"
Const MSG34 = "Imposible cerrar con transacciones pendientes"
Const MSG35 = "Debe cerrar primero"
Const MSG36 = "Abrir base de datos de Microsoft Access"
Const MSG37 = "Abrir base de datos de dBase"
Const MSG38 = "Abrir base de datos de FoxPro"
Const MSG39 = "Abrir base de datos de Paradox"
Const MSG40 = "Abrir archivo de Excel"
Const MSG41 = "Abrir base de datos de Btrieve"
Const MSG42 = "Abrir base de datos de texto"
Const MSG43 = "Abriendo base de datos"
Const MSG44 = "NOTA: Es recomendable usar tablas adjuntas"
Const MSG45 = "Reparando "
Const MSG46 = "¿Desea intentar repararlo?"
Const MSG47 = "Escriba un nombre de directorio para una nueva base de datos ISAM:"
Const MSG48 = "Seleccione la base de datos de Microsoft Access para compactar"
Const MSG49 = "MDB de Microsoft Access (*.mdb)|*.mdb"
Const MSG50 = "|Todos los archivos (*.*)|*.*"
Const MSG51 = "Seleccione la base de datos de Microsoft Access a compactar"
Const MSG52 = "¿Desea codificar la base de datos compactada?"
Const MSG53 = "Compactando "
Const MSG54 = "¿Desea abrir la base de datos compactada recientemente?"
Const MSG55 = "Seleccione una base de datos de Microsoft Access para crear"
Const MSG56 = "Exportando tabla: "
Const MSG57 = "Exportar "
Const MSG58 = "en "
Const MSG59 = "Creando índices:"
Const MSG60 = "Exportado correctamente:"
Const MSG61 = "Instrucción SQL exportada correctamente."
Const MSG62 = "La tabla ya existe. ¿Desea sobrescribirla?"
Const MSG63 = "Importando tabla: "
Const MSG64 = "Importado correctamente:"
Const MSG65 = "Nombre no válido de directorio"
'>>>>>>>>>>>>>>>>>>>>>>>>


'declaraciones del API
Declare Function OSGetPrivateProfileString% Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
Declare Function OSWritePrivateProfileString% Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Declare Function SQLDataSources% Lib "ODBC32.DLL" (ByVal henv&, ByVal fdir%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDesc$, ByVal cbDescMax%, pcbDesc%)
Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Declare Function GetDesktopWindow Lib "user32" () As Long

'variables objeto globales
Global gVDClass      As New VisDataClass
Global gnodDBNode    As Node        'nodo de base de datos actual
Global gnodDBNode2   As Node        'copia de seguridad del nodo de la base de datos actual
Global gwsMainWS     As Workspace   'objeto entorno de trabajo principal
Global gdbCurrentDB  As Database    'objeto base de datos principal
Global gbDBOpenFlag  As Integer     'indicador para saber si una base de datos está abierta
Global gPropObject   As Object      'objeto para mostrar propiedades
Global gDataCtlObj   As Object      'objeto control Data global
Global gtdfTableDef  As TableDef    'Tabledef global utilizada por frmTblStruct
Global gnFormType    As Integer     'tipo de formulario elegido en el formulario principal
                                    '0 = control Data
                                    '1 = sin control Data
                                    '2 = control Grid
Global gnRSType      As Integer     'tipo de Recordset elegido en el formulario principal
                                    '0 = tabla
                                    '1 = dinámico
                                    '2 = instantánea

'variables de base de datos globales
Global gsDataType       As String   'datos finales = cadena de conexión
                                    'para todas menos para Access
Global gsDBName         As String   'nombre de la base de datos actual
Global gsODBCDatasource As String   'valores ODBC globales
Global gsODBCDatabase   As String   '       "
Global gsODBCUserName   As String   '       "
Global gsODBCPassword   As String   '       "
Global gsODBCDriver     As String   '       "
Global gsODBCServer     As String   '       "
Global gsTblName        As String   '
Global glQueryTimeout   As Long     '
Global glLoginTimeout   As Long     '
Global gsTableDynaFilter As String  '
Global gnReadOnly       As Integer  'indicador de sólo lectura

'otras variables globales
Global gsZoomData       As String   'pasa información al formulario Zoom

'variables multiusuario
Global gnMURetryCnt     As Integer
Global gnMUDelay        As Integer
Global gnMULocking      As Integer  'indicador de bloqueo pesimista u optimista

'valores globales de búsqueda utilizados para pasar información entre
'el formulario Dynaset y el cuadro de diálogo de búsqueda
Global gbFindFailed     As Boolean
Global gsFindExpr       As String
Global gsFindOp         As String
Global gsFindField      As String
Global gnFindType       As Integer
Global gbFromTableView  As Boolean

'valores globales de búsqueda utilizados para pasar información entre
'el formulario Table y el cuadro de diálogo de búsqueda
Global gsSeekOperator   As String
Global gsSeekValue      As String

'indicadores globales
Global gbDBChanged      As Boolean   '
Global gbTransPending   As Boolean   'utilizado para la administración de transacciones
Global gbFromSQL        As Boolean   'el origen de la instrucción fue el formulario SQL
Global gbAddTableFlag   As Boolean   'designa nuevo o diseño
Global gbSettingDataCtl As Boolean   'utilizado para restablecer las propiedades del control Data

'variables globales utilizadas en el código de Importar/Exportar
Global gnDataType As Integer
Global gImpDB As Database
Global gExpDB As Database
Global gExpTable As String

'tipo de datos finales utilizados en las cadenas de conexión
Global Const gsMSACCESS = "Microsoft Access"
Global Const gsDBASEIII = "Dbase III;"
Global Const gsDBASEIV = "Dbase IV;"
Global Const gsDBASE5 = "Dbase 5.0;"
Global Const gsFOXPRO20 = "FoxPro 2.0;"
Global Const gsFOXPRO25 = "FoxPro 2.5;"
Global Const gsFOXPRO26 = "FoxPro 2.6;"
Global Const gsFOXPRO30 = "FoxPro 3.0;"
Global Const gsPARADOX3X = "Paradox 3.X;"
Global Const gsPARADOX4X = "Paradox 4.X;"
Global Const gsPARADOX5X = "Paradox 5.X;"
Global Const gsBTRIEVE = "Btrieve;"
Global Const gsEXCEL30 = "Excel 3.0;"
Global Const gsEXCEL40 = "Excel 4.0;"
Global Const gsEXCEL50 = "Excel 5.0;"
Global Const gsTEXTFILES = "Text;"
Global Const gsSQLDB = "ODBC;"

'tipos de datos de exportar/importar
Global Const gnDT_NONE = -1
Global Const gnDT_MSACCESS = 0
Global Const gnDT_DBASEIV = 1
Global Const gnDT_DBASEIII = 2
Global Const gnDT_FOXPRO26 = 3
Global Const gnDT_FOXPRO25 = 4
Global Const gnDT_FOXPRO20 = 5
Global Const gnDT_PARADOX4X = 6
Global Const gnDT_PARADOX3X = 7
Global Const gnDT_BTRIEVE = 8
Global Const gnDT_EXCEL50 = 9
Global Const gnDT_EXCEL40 = 10
Global Const gnDT_EXCEL30 = 11
Global Const gnDT_TEXTFILE = 12
Global Const gnDT_SQLDB = 13

'constantes globales
Global Const gsDEFAULT_DRIVER = "SQL Server"  'utilizado en Registerdatabase
Global Const gnEOF_ERR = 626                  '
Global Const gnFTBLS = 0                      '
Global Const gnFFLDS = 1                      '
Global Const gnFINDX = 2                      '
Global Const gnMAX_GRID_ROWS = 31999          '
Global Const gnMAX_MEMO_SIZE = 20000          '
Global Const gnGETCHUNK_CUTOFF = 50           '

Global Const gnFORM_DATACTL = 0               '
Global Const gnFORM_NODATACTL = 1             '
Global Const gnFORM_DATAGRID = 2              '

Global Const gnRS_TABLE = vbRSTypeTable
Global Const gnRS_DYNASET = vbRSTypeDynaset
Global Const gnRS_SNAPSHOT = vbRSTypeSnapShot
Global Const gnRS_PASSTHRU = 8

Global Const gnCTLARRAYHEIGHT = 340&          '
Global Const gnSCREEN = 0                     'utilizado para centrar formularios en la pantalla
Global Const gnMDIFORM = 1                    'utilizado para centrar formularios en frmMDI

Global Const TABLE_STR = "Table"
Global Const ATTACHED_STR = "Attached"
Global Const QUERY_STR = "Query"
Global Const FIELD_STR = "Field"
Global Const FIELDS_STR = "Fields"
Global Const INDEX_STR = "Index"
Global Const INDEXES_STR = "Indexes"
Global Const PROPERTY_STR = "Property"
Global Const PROPERTIES_STR = "Properties"

Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"

Sub Main()
  frmMDI.Show
End Sub


'------------------------------------------------------------
'esta función devuelve el tipo del Querydef
'del elemento seleccionado en la lista de Querydefs
'del formulario frmTables
'------------------------------------------------------------
Function ActionQueryType(qdf As QueryDef) As String
  
  'comprueba si es una consulta de acciones
  If (qdf.Type And dbQAction) = 0 Then
    ActionQueryType = vbNullString
    Exit Function
  End If
  
  'tiene que ser una consulta de acciones
  Select Case qdf.Type
    Case dbQCrosstab
      ActionQueryType = "Cross Tab"
    Case dbQDelete
      ActionQueryType = "Delete"
    Case dbQUpdate
      ActionQueryType = "Update"
    Case dbQAppend
      ActionQueryType = "Append"
    Case dbQMakeTable
      ActionQueryType = "Make Table"
    Case dbQDDL
      ActionQueryType = "DDL"
    Case dbQSQLPassThrough
      ActionQueryType = "SQLPassThrough"
    Case dbQSetOperation
      ActionQueryType = "Set Operation"
    Case dbQSPTBulk
      ActionQueryType = "SPT Bulk"
    Case Else
      ActionQueryType = vbNullString
  End Select

End Function

'-------------------------------------------------------------
'esta función agrega [] a los nombres de objeto que lo
'necesiten porque contienen espacios
'-------------------------------------------------------------
Function AddBrackets(rObjName As String) As String
  'agrega corchetes a los nombres de objetos que contengan espacios
  If InStr(rObjName, " ") > 0 And Mid(rObjName, 1, 1) <> "[" Then
    AddBrackets = "[" & rObjName & "]"
  Else
    AddBrackets = rObjName
  End If
End Function

'------------------------------------------------------------
'esta función comprueba si hay una transacción pendiente
'y presenta un mensaje
'------------------------------------------------------------
Function CheckTransPending(MSG As String) As Integer

  If gbTransPending Then
    MsgBox MSG & vbCrLf & MSG1, 48
    CheckTransPending = True
  Else
    CheckTransPending = False
  End If

End Function

'------------------------------------------------------------
'borra los campos de datos en los formularios de tablas y de
'conjuntos de registros
'------------------------------------------------------------
Sub ClearDataFields(frm As Form, nCnt As Integer)
  Dim i As Integer

  'borra los campos del formulario principal
  For i = 0 To nCnt - 1
    frm.txtFieldData(i).Text = vbNullString
  Next
End Sub

'------------------------------------------------------------
'esta subrutina cierra todos los formularios frmDynaSnap,
'frmTableObj y frmDataGrid buscando formularios con un
'Tag establecido a "Recordset"
'------------------------------------------------------------
Sub CloseAllRecordsets()
  Dim i As Integer

  MsgBar MSG2, True
  While i < Forms.Count
    If Forms(i).Tag = "Recordset" Then
      Unload Forms(i)
    Else
      i = i + 1
    End If
  Wend
  MsgBar vbNullString, False

End Sub

'------------------------------------------------------------
'esta función copia datos entre tablas
'desde el formulario frmCopyStruct
'Demuestra el uso de las transacciones para acelerar este
'tipo de operación
'------------------------------------------------------------
Function CopyData(rFromDB As Database, rToDB As Database, rFromName As String, rToName As String) As Integer
  On Error GoTo CopyErr

  Dim recRecordset1 As Recordset, recRecordset2 As Recordset
  Dim i As Integer
  Dim nRC As Integer
  Dim fld As Field

  'abre ambos conjuntos de registros
  Set recRecordset1 = rFromDB.OpenRecordset(rFromName)
  Set recRecordset2 = rToDB.OpenRecordset(rToName)
  gwsMainWS.BeginTrans
  While recRecordset1.EOF = False
    recRecordset2.AddNew
    'este bucle copia los datos de cada campo
    'a la nueva tabla
'    For Each fld In recRecordset1.Fields
    For i = 0 To recRecordset1.Fields.Count - 1
      Set fld = recRecordset1.Fields(i)
      recRecordset2(fld.Name).Value = fld.Value
    Next
    recRecordset2.Update
    recRecordset1.MoveNext
    nRC = nRC + 1
    'esta comprobación confirmará las transacciones cada 1000 registros
    If nRC = 1000 Then
      gwsMainWS.CommitTrans
      gwsMainWS.BeginTrans
      nRC = 0
    End If
  Wend
  gwsMainWS.CommitTrans

  CopyData = True
  Exit Function

CopyErr:
  gwsMainWS.Rollback
  ShowError
  CopyData = False
End Function

'------------------------------------------------------------
'esta función copia la estructura de una tabla en una
'nueva tabla en la misma base de datos o en otra diferente
'------------------------------------------------------------
Function CopyStruct(vFromDB As Database, vToDB As Database, vFromName As String, vToName As String, bCreateIndex As Integer) As Integer
  On Error GoTo CSErr

  Dim i As Integer
  Dim tblTableDefObj As TableDef
  Dim fldFieldObj As Field
  Dim indIndexObj As Index
  Dim tdf As TableDef
  Dim fld As Field
  Dim idx As Index
  
  'busca para ver si la tabla existe
NameSearch:
'  For Each tdf In vToDB.Tabledefs
  For i = 0 To vToDB.TableDefs.Count - 1
    Set tdf = vToDB.TableDefs(i)
    If UCase(tdf.Name) = UCase(vToName) Then
      If MsgBox(MSG3, 4) = vbYes Then
         vToDB.TableDefs.Delete tdf.Name
      Else
         vToName = InputBox(MSG4)
         If Len(vToName) = 0 Then
           Exit Function
         Else
           GoTo NameSearch
         End If
      End If
      Exit For
    End If
  Next
  
  Set tblTableDefObj = gdbCurrentDB.CreateTableDef()
    
  'quita el propietario si fuera necesario
  tblTableDefObj.Name = StripOwner(vToName)

  'crea los campos
'  For Each fld In vFromDB.Tabledefs(vFromName).Fields
  For i = 0 To vFromDB.TableDefs(vFromName).Fields.Count - 1
    Set fld = vFromDB.TableDefs(vFromName).Fields(i)
    Set fldFieldObj = vFromDB.TableDefs(vFromName).CreateField(fld.Name, fld.Type, fld.Size)
    tblTableDefObj.Fields.Append fldFieldObj
  Next

  'crea los índices
  If bCreateIndex <> False Then
'    For Each idx In vFromDB.Tabledefs(vFromName).Indexes
    For i = 0 To vFromDB.TableDefs(vFromName).Indexes.Count - 1
      Set idx = vFromDB.TableDefs(vFromName).Indexes(i)
      Set indIndexObj = vFromDB.TableDefs(vFromName).CreateIndex(idx.Name)
      With indIndexObj
        indIndexObj.Fields = idx.Fields
        indIndexObj.Unique = idx.Unique
        If gsDataType <> gsSQLDB Then
          indIndexObj.Primary = idx.Primary
        End If
      End With
      tblTableDefObj.Indexes.Append indIndexObj
    Next
  End If

  'agrega la nueva tabla
  vToDB.TableDefs.Append tblTableDefObj

  CopyStruct = True
  Exit Function

CSErr:
  ShowError
  CopyStruct = False
End Function

'------------------------------------------------------------
'esta función llena un cuadro de lista o combinado con las
'tablas (y Querydefs) del formulario Tables
'ItemData se establece a 0 para las Tabledef y a 1 para las
'Querydef
'------------------------------------------------------------
Sub GetTableList(rctl As Control, rbIncludeQDFs As Integer, rbIncludeSys As Integer, rbStripConnect As Integer)
  On Error GoTo FTLErr
  
  Dim i As Integer
  Dim sTmp As String
  Dim tbl As TableDef
  Dim qdf As QueryDef
  
  'agrega las Tabledefs
  For Each tbl In gdbCurrentDB.TableDefs
    sTmp = tbl.Name
    If rbIncludeSys Then
      rctl.AddItem sTmp
      rctl.ItemData(rctl.NewIndex) = 0
    Else
      If (gdbCurrentDB.TableDefs(sTmp).Attributes And dbSystemObject) = 0 Then
        rctl.AddItem sTmp
        rctl.ItemData(rctl.NewIndex) = 0
      End If
    End If
  Next
  
  'agrega las Querydefs
  If rbIncludeQDFs Then
    For Each qdf In gdbCurrentDB.QueryDefs
      rctl.AddItem qdf.Name
      rctl.ItemData(rctl.NewIndex) = 1
    Next
  End If
  
  Exit Sub
  
FTLErr:
  ShowError
End Sub

'------------------------------------------------------------
'esta función devuelve el tipo de campo numérico
'de la cadena pasada
'------------------------------------------------------------
Function GetFieldType(rFldType As String) As Integer
  'devuelve la longitud del campo
  If rFldType = "Text" Then
    GetFieldType = dbText
  Else
    Select Case rFldType
      Case "Counter"
        GetFieldType = dbLong
      Case "Boolean"
        GetFieldType = dbBoolean
      Case "Byte"
        GetFieldType = dbByte
      Case "Integer"
        GetFieldType = dbInteger
      Case "Long"
        GetFieldType = dbLong
      Case "Currency"
        GetFieldType = dbCurrency
      Case "Single"
        GetFieldType = dbSingle
      Case "Double"
        GetFieldType = dbDouble
      Case "Date/Time"
        GetFieldType = dbDate
      Case "Binary"
        GetFieldType = dbLongBinary
      Case "Memo"
        GetFieldType = dbMemo
    End Select
  End If

End Function

'------------------------------------------------------------
'esta función devuelve la anchura apropiada para el
'tipo de campo pasado, para ser utilizada como anchura del
'control en los formularios frmDynaSnap y frmTableObj
'------------------------------------------------------------
Function GetFieldWidth(rType As Integer)
  Select Case rType
    Case dbBoolean
      GetFieldWidth = 850
    Case dbByte
      GetFieldWidth = 650
    Case dbInteger
      GetFieldWidth = 900
    Case dbLong
      GetFieldWidth = 1100
    Case dbCurrency
      GetFieldWidth = 1800
    Case dbSingle
      GetFieldWidth = 1800
    Case dbDouble
      GetFieldWidth = 2200
    Case dbDate
      GetFieldWidth = 2000
    Case dbText
      GetFieldWidth = 3250
    Case dbLongBinary
      GetFieldWidth = 3250
    Case dbMemo
      GetFieldWidth = 3250
    Case Else
      GetFieldWidth = 3250
  End Select

End Function

'------------------------------------------------------------
'esta función devuelve el valor del  archivo INI para el
'elemento y sección pasados
'------------------------------------------------------------
Function GetINIString(ByVal vsItem As String, ByVal vsDefault As String) As String
  GetINIString = GetSetting(APP_CATEGORY, App.Title, vsItem, vsDefault)
End Function

'------------------------------------------------------------
'esta subrutina oculta los menús y las barras de herramientas
'que sólo se aplican cuando hay una base de datos abierta
'------------------------------------------------------------
Sub HideDBTools()
  frmMDI.mnuDBClose.Enabled = False
  frmMDI.mnuDBImpExp.Enabled = False
  frmMDI.mnuUtil.Enabled = False
  frmMDI.mnuUBar1.Visible = False
  frmMDI.mnuUAttachments.Visible = False
  frmMDI.mnuUGroupsUsers.Visible = False
  frmMDI.mnuUSystemDB.Visible = False
  frmMDI.tlbToolBar.Buttons("BeginTrans").Enabled = False
  frmMDI.tlbToolBar.Buttons("Commit").Enabled = False
  frmMDI.tlbToolBar.Buttons("Rollback").Enabled = False
End Sub

'------------------------------------------------------------
'esta subrutina presenta el mensaje pasado en la barra de
'estado de la parte inferior del formulario MDI
'------------------------------------------------------------
Sub MsgBar(rsMsg As String, rPauseFlag As Integer)
  If Len(rsMsg) = 0 Then
    Screen.MousePointer = vbDefault
    frmMDI.stsStatusBar.Panels(1).Text = MSG5
  Else
    If rPauseFlag Then
      frmMDI.stsStatusBar.Panels(1).Text = rsMsg & MSG6
    Else
      frmMDI.stsStatusBar.Panels(1).Text = rsMsg
    End If
  End If
  frmMDI.stsStatusBar.Refresh
End Sub

'==================================================
' Subrutina: ObjectExists
'
' Propósito: Determina si existe o no un miembro
'          igual que MemberExists excepto en que el primer
'          argumento es un objeto para poder pasar colecciones
'          como VBComponents, VBProjects, etc.
' Argumentos:
'   pColl: Nombre de la colección a comprobar
'   sMemName: Nombre (clave) del miembro a comprobar
' Salida:
'   True: el miembro existe en la colección
'   False: el miembro no existe en la colección
' Mantenimiento: J$
'==================================================
Function ObjectExists(pColl As Object, sMemName As String) As Boolean
Dim pObj As Object
  
  On Error Resume Next
  Err = 0
  Set pObj = pColl(sMemName)
  ObjectExists = (Err = 0)
End Function


'------------------------------------------------------------
'esta subrutina actualiza cualquier lista de tablas
'pasada como objeto
'------------------------------------------------------------
Sub RefreshTables(rListObject As Object)
  On Error GoTo TRefErr

  Dim tdf As TableDef
  Dim qdf As QueryDef
  Dim sTmp As String
    
  Dim i As Integer
    
  MsgBar MSG7, True
  Screen.MousePointer = vbHourglass

  'si se le llama para actualizar la ventana de base de datos,
  'ignora el método antiguo de llenar
  'un cuadro de lista con los nombres de las tablas
  If rListObject Is Nothing Then GoTo LoadTreeView

  rListObject.Clear
  If frmMDI.mnuPAllowSys.Checked Then
    'enumera todas las tablas
    For Each tdf In gdbCurrentDB.TableDefs
      If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
        If Left(tdf.Connect, 1) = ";" Then
          'tiene que ser una tabla adjunta de Microsoft Access
          rListObject.AddItem tdf.Name & " -> Microsoft Access"
        Else
          'tiene que ser una tabla adjunta ISAM
          rListObject.AddItem tdf.Name & " -> " & Left(tdf.Connect, InStr(tdf.Connect, ";") - 1)
        End If
      ElseIf (tdf.Attributes And dbAttachedODBC) = dbAttachedODBC Then
        rListObject.AddItem tdf.Name & " -> ODBC"
      Else
        rListObject.AddItem tdf.Name
      End If
    Next
  Else
    'no enumera las tablas del sistema
    For Each tdf In gdbCurrentDB.TableDefs
      If (tdf.Attributes And dbSystemObject) = 0 Then
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
          If Left(tdf.Connect, 1) = ";" Then
            'tiene que ser una tabla adjunta de Microsoft Access
            rListObject.AddItem tdf.Name & " -> Microsoft Access"
          Else
            'tiene que ser una tabla adjunta ISAM
            rListObject.AddItem tdf.Name & " -> " & Left(tdf.Connect, InStr(tdf.Connect, ";") - 1)
          End If
        ElseIf (tdf.Attributes And dbAttachedODBC) = dbAttachedODBC Then
          rListObject.AddItem tdf.Name & " -> ODBC"
        Else
          rListObject.AddItem tdf.Name
        End If
      End If
    Next
  End If
  'selecciona el primer elemento si lo hay
  If rListObject.ListCount > 0 Then
    rListObject.ListIndex = 0
  End If
  
LoadTreeView:
  frmDatabase.LoadDatabase
  
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  Exit Sub

TRefErr:
  ShowError
End Sub

'------------------------------------------------------------
'esta función devuelve el tamaño del tipo de campo
'pasado para ser utilizado en el formulario frmAddField
'------------------------------------------------------------
Function SetFldProperties(rnType As Integer) As Integer
  'devuelve la longitud del campo
  Select Case rnType
    Case dbBoolean
      SetFldProperties = 1
    Case dbByte
      SetFldProperties = 1
    Case dbInteger
      SetFldProperties = 2
    Case dbLong
      SetFldProperties = 4
    Case dbCurrency
      SetFldProperties = 8
    Case dbSingle
      SetFldProperties = 4
    Case dbDouble
      SetFldProperties = 8
    Case dbDate
      SetFldProperties = 8
    Case dbText
      SetFldProperties = 50
    Case dbLongBinary
      SetFldProperties = 0
    Case dbMemo
      SetFldProperties = 0
  End Select
End Function

'------------------------------------------------------------
'esta subrutina muestra los menús y las barras de herramientas
'que sólo se aplican cuando hay una base de datos abierta
'------------------------------------------------------------
Sub ShowDBTools()
  Dim sTmp As String

  frmMDI.mnuDBClose.Enabled = True
  frmMDI.mnuDBImpExp.Enabled = True
  frmMDI.mnuUtil.Enabled = True
  frmMDI.tlbToolBar.Buttons("BeginTrans").Enabled = True
  frmMDI.tlbToolBar.Buttons("Commit").Enabled = False
  frmMDI.tlbToolBar.Buttons("Rollback").Enabled = False
  frmMDI.tlbToolBar.Refresh

  'establece elementos generales que se aplican o no a las MDBs
  If gsDataType = gsMSACCESS Then
    frmMDI.mnuUBar1.Visible = True
    frmMDI.mnuUAttachments.Visible = True
    frmMDI.mnuUGroupsUsers.Visible = True
    frmMDI.mnuUSystemDB.Visible = True
    frmSQL.cmdSaveQueryDef.Visible = True
    frmMDI.mnuDBPURename.Visible = True
  Else
    frmSQL.cmdSaveQueryDef.Visible = False
    frmMDI.mnuDBPURename.Visible = False
  End If

  'establece los elementos específicos de ODBC
  If gsDataType = gsSQLDB Then
    If gnRSType = gnRS_TABLE Then
      frmMDI.tlbToolBar.Buttons("Dynaset").Value = tbrPressed
      gnRSType = gnRS_DYNASET
    End If
    frmMDI.tlbToolBar.Buttons("PassThrough").Visible = True
    frmMDI.tlbToolBar.Buttons("Table").Visible = False
  Else
    If gnRSType = gnRS_PASSTHRU Then
      frmMDI.tlbToolBar.Buttons("Dynaset").Value = tbrPressed
      gnRSType = gnRS_DYNASET
    End If
    frmMDI.tlbToolBar.Buttons("PassThrough").Visible = False
    frmMDI.tlbToolBar.Buttons("Table").Visible = True
  End If
  frmMDI.tlbToolBar.Refresh
  'muestra 2 formularios hijos principales
  frmDatabase.Show
  frmSQL.Show
End Sub

'------------------------------------------------------------
'esta subrutina presenta el mensaje de error con su código de
'error y solicita que se muestre la colección Errors si se
'trata de un error de acceso a datos
'------------------------------------------------------------
Sub ShowError()
  Dim sTmp As String

  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False

  sTmp = "Ocurrió el siguiente error:" & vbCrLf & vbCrLf
  'agrega la cadena del error
  sTmp = sTmp & Err.Description & vbCrLf
  'agrega el número de error
  sTmp = sTmp & MSG8 & Err
  
  Beep
  'comprueba si el error viene de la colección db errors
  If DBEngine.Errors.Count > 0 Then
    If DBEngine.Errors(0).Number = Err Then
      'agrega la pregunta sobre la presentación de la colección Errors
      sTmp = sTmp & vbCrLf & vbCrLf & MSG9
      'emite un pitido y muestra el error
      If MsgBox(sTmp, vbYesNo + vbQuestion) = vbYes Then
        RefreshErrors
      End If
    Else
      MsgBox sTmp
    End If
  Else
    MsgBox sTmp
  End If

End Sub

'------------------------------------------------------------
'esta función elimina la cadena de conexión de la tabla adjunta
'------------------------------------------------------------
Function StripConnect(rsTblName As String) As String
  If InStr(rsTblName, "->") > 0 Then
    StripConnect = Left(rsTblName, InStr(rsTblName, "->") - 2)
  Else
    StripConnect = rsTblName
  End If

End Function

'------------------------------------------------------------
'esta función elimina los corchetes ([]), de los objetos de datos
'------------------------------------------------------------
Function StripBrackets(rsObjName As String) As String
  'agrega corchetes a los nombres de objeto que contengan espacios
  If Mid(rsObjName, 1, 1) = "[" Then
    StripBrackets = Mid(rsObjName, 2, Len(rsObjName) - 2)
  Else
    StripBrackets = rsObjName
  End If

End Function

'------------------------------------------------------------
'esta función elimina el nombre del archivo de una
'cadena ruta\archivo
'------------------------------------------------------------
Function StripFileName(rsFileName As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = "\" Then
      Exit For
    End If
  Next

  StripFileName = Mid(rsFileName, 1, i - 1)

End Function

'------------------------------------------------------------
'esta función elimina los caracteres no ASCII del campo memo
'antes de presentarlo (no se asegura que sea siempre necesario)
'------------------------------------------------------------
Function StripNonAscii(rvntVal As Variant) As String
  Dim i As Integer
  Dim sTmp As String

  'se termina para permitir los caracteres DBCS
  StripNonAscii = rvntVal
  Exit Function

  For i = 1 To Len(rvntVal)
    If Asc(Mid(rvntVal, i, 1)) < 32 Or Asc(Mid(rvntVal, i, 1)) > 126 Then
      sTmp = sTmp & " "
    Else
      sTmp = sTmp & Mid(rvntVal, i, 1)
    End If
  Next

  StripNonAscii = sTmp

End Function

'------------------------------------------------------------
'elimina el propietario de los nombres de tablas ODBC
'------------------------------------------------------------
Function StripOwner(rsTblName As String) As String

  If InStr(rsTblName, ".") > 0 Then
    rsTblName = Mid(rsTblName, InStr(rsTblName, ".") + 1, Len(rsTblName))
  End If
  StripOwner = rsTblName

End Function

'------------------------------------------------------------
'devuelve la cadena True o False
'------------------------------------------------------------
Function stTrueFalse(rvntTF As Variant) As String
  If rvntTF Then
    stTrueFalse = "True"
  Else
    stTrueFalse = "False"
  End If
End Function

'------------------------------------------------------------
'devuelve "" si el campo es nulo
'------------------------------------------------------------
Function vFieldVal(rvntFieldVal As Variant) As Variant
  If IsNull(rvntFieldVal) Then
    vFieldVal = vbNullString
  Else
    vFieldVal = CStr(rvntFieldVal)
  End If
End Function

'------------------------------------------------------------
'carga toda la configuración de VisData almacenada en
'el archivo INI
'------------------------------------------------------------
Sub LoadINISettings()
  On Error Resume Next
  
  Dim sTmp As String
  Dim x As Integer

  glQueryTimeout = Val(GetINIString("QueryTimeout", "5"))
  glLoginTimeout = Val(GetINIString("LoginTimeout", "20"))
  
  
  frmMDI.mnuPOpenOnStartup.Checked = Val(GetINIString("OpenOnStartup", "0"))
  frmMDI.mnuPAllowSys.Checked = Val(GetINIString("AllowSys", "0"))

  'obtiene las bases de datos más recientemente utilizadas
  For x = 1 To 8
    sTmp = GetINIString("MRUDatabase" & x, "")
    If Len(sTmp) > 0 Then
      frmMDI.mnuBarMRU.Visible = True
      frmMDI.mnuDBMRU(x).Caption = "&" & x & " " & sTmp
      frmMDI.mnuDBMRU(x).Visible = True
      sTmp = GetINIString("MRUConnect" & x, "")
      frmMDI.mnuDBMRU(x).Tag = sTmp
    End If
  Next

  'obtiene la base de datos más recientemente utilizada desde el archivo INI
  gsDataType = GetINIString("DataType", vbNullString)
  gsDBName = GetINIString("DatabaseName", vbNullString)
  gsODBCDatasource = GetINIString("ODBCDatasource", vbNullString)
  gsODBCDatabase = GetINIString("ODBCDatabase", vbNullString)
  gsODBCUserName = GetINIString("ODBCUserName", vbNullString)
  gsODBCPassword = GetINIString("ODBCPassword", vbNullString)
  gsODBCDriver = GetINIString("ODBCDriver", vbNullString)
  gsODBCServer = GetINIString("ODBCServer", vbNullString)

  sTmp = GetINIString("ViewMode", CStr(gnFORM_NODATACTL))
  Select Case Val(sTmp)
    Case gnFORM_NODATACTL
      gnFormType = gnFORM_NODATACTL
    Case gnFORM_DATACTL
      gnFormType = gnFORM_DATACTL
    Case gnFORM_DATAGRID
      gnFormType = gnFORM_DATAGRID
  End Select
  sTmp = GetINIString("RecordsetType", CStr(vbRSTypeDynaset))
  Select Case Val(sTmp)
    Case vbRSTypeTable
      gnRSType = gnRS_TABLE
    Case vbRSTypeDynaset
      gnRSType = gnRS_DYNASET
    Case vbRSTypeSnapShot
      gnRSType = gnRS_SNAPSHOT
    Case gnRS_PASSTHRU
      gnRSType = gnRS_PASSTHRU
  End Select
  
  DoEvents
  Select Case gnFormType
    Case gnFORM_NODATACTL
      frmMDI.tlbToolBar.Buttons("NoDataControl").Value = tbrPressed
    Case gnFORM_DATACTL
      frmMDI.tlbToolBar.Buttons("DataControl").Value = tbrPressed
    Case gnFORM_DATAGRID
      frmMDI.tlbToolBar.Buttons("DBGrid").Value = tbrPressed
  End Select
  Select Case gnRSType
    Case vbRSTypeDynaset
      frmMDI.tlbToolBar.Buttons("Dynaset").Value = tbrPressed
    Case vbRSTypeSnapShot
      frmMDI.tlbToolBar.Buttons("Snapshot").Value = tbrPressed
    Case vbRSTypeTable
      frmMDI.tlbToolBar.Buttons("Table").Value = tbrPressed
    Case gnRS_PASSTHRU
      frmMDI.tlbToolBar.Buttons("PassThrough").Value = tbrPressed
  End Select
  
End Sub

'------------------------------------------------------------
'guarda los valores actuales de VisData en VISDATA.INI
'------------------------------------------------------------
Sub SaveINISettings()
  On Error Resume Next

  Dim i As Integer
  
  SaveSetting APP_CATEGORY, App.Title, "DataType", gsDataType
  SaveSetting APP_CATEGORY, App.Title, "DatabaseName", gsDBName
  SaveSetting APP_CATEGORY, App.Title, "ODBCDatasource", gsODBCDatasource
  SaveSetting APP_CATEGORY, App.Title, "ODBCDatabase", gsODBCDatabase
  SaveSetting APP_CATEGORY, App.Title, "ODBCUserName", gsODBCUserName
  SaveSetting APP_CATEGORY, App.Title, "ODBCPassword", gsODBCPassword
  SaveSetting APP_CATEGORY, App.Title, "ODBCDriver", gsODBCDriver
  SaveSetting APP_CATEGORY, App.Title, "ODBCServer", gsODBCServer
  SaveSetting APP_CATEGORY, App.Title, "QueryTimeout", glQueryTimeout
  SaveSetting APP_CATEGORY, App.Title, "LoginTimeout", glLoginTimeout
  DBEngine.LoginTimeout = glLoginTimeout
  SaveSetting APP_CATEGORY, App.Title, "ViewMode", gnFormType
  SaveSetting APP_CATEGORY, App.Title, "RecordsetType", gnRSType
  
  SaveSetting APP_CATEGORY, App.Title, "OpenOnStartup", IIf(frmMDI.mnuPOpenOnStartup.Checked, "-1", "0")
  SaveSetting APP_CATEGORY, App.Title, "AllowSys", IIf(frmMDI.mnuPAllowSys.Checked, "-1", "0")

  For i = 1 To 8
    If frmMDI.mnuDBMRU(i).Visible Then
      SaveSetting APP_CATEGORY, App.Title, "MRUDatabase" & i, Mid(frmMDI.mnuDBMRU(i).Caption, 4, Len(frmMDI.mnuDBMRU(i).Caption))
      SaveSetting APP_CATEGORY, App.Title, "MRUConnect" & i, frmMDI.mnuDBMRU(i).Tag
    Else
      SaveSetting APP_CATEGORY, App.Title, "MRUDatabase" & i, ""
      SaveSetting APP_CATEGORY, App.Title, "MRUConnect" & i, ""
    End If
  Next

  SaveSetting APP_CATEGORY, App.Title, "WindowState", frmMDI.WindowState
  If frmMDI.WindowState = vbNormal Then
    SaveSetting APP_CATEGORY, App.Title, "WindowTop", frmMDI.Top
    SaveSetting APP_CATEGORY, App.Title, "WindowLeft", frmMDI.Left
    SaveSetting APP_CATEGORY, App.Title, "WindowWidth", frmMDI.Width
    SaveSetting APP_CATEGORY, App.Title, "WindowHeight", frmMDI.Height
  End If
  SaveSetting APP_CATEGORY, App.Title, "ViewMode", gnFormType
  SaveSetting APP_CATEGORY, App.Title, "RecordsetType", gnRSType

End Sub

'------------------------------------------------------------
'esta subrutina abrirá el formulario del tipo de datos
'apropiado y presenta el mensaje adecuado en la barra de
'estado de acuerdo con las opciones seleccionadas por el
'usuario en el formulario MDI principal
'------------------------------------------------------------
Sub OpenTable(rName As String)
  On Error GoTo OpenTableErr

  Dim rsTmp As Recordset
  Dim sTmp As String
  Dim nAttach As Integer
  Dim frmTmp As Form
  
  If gsDataType = gsMSACCESS Then   'busca tablas adjuntas si es una MDB
    If (gdbCurrentDB.TableDefs(rName).Attributes And dbAttachedTable) = dbAttachedTable Then
      nAttach = 1
    ElseIf (gdbCurrentDB.TableDefs(rName).Attributes And dbAttachedODBC) = dbAttachedODBC Then
      nAttach = 2
    End If
    If nAttach > 0 And gnRSType = gnRS_TABLE Then
      Beep
      If MsgBox(MSG10, vbYesNo + vbQuestion) = vbYes Then
        frmMDI.tlbToolBar.Buttons("Dynaset").Value = tbrPressed         'restablece a Dynaset
      Else
        Exit Sub
      End If
    End If
  End If
  
  If nAttach > 0 Then
    If gnRSType = gnRS_DYNASET Then
      sTmp = MSG11
    ElseIf gnRSType = gnRS_SNAPSHOT Then
      sTmp = MSG12
    End If
  Else
    If gnRSType = gnRS_TABLE Then
      sTmp = MSG13
    ElseIf gnRSType = gnRS_DYNASET Then
      sTmp = MSG14
    ElseIf gnRSType = gnRS_SNAPSHOT Then
      sTmp = MSG15
    ElseIf gnRSType = gnRS_PASSTHRU Then
      sTmp = MSG16
    End If
  End If
  
  MsgBar sTmp, True
  
  Screen.MousePointer = vbHourglass
  If gnRSType = gnRS_TABLE Then
    Set rsTmp = gdbCurrentDB.OpenRecordset(rName, dbOpenTable)
    sTmp = "Tabla:"
  ElseIf gnRSType = gnRS_DYNASET Then
    Set rsTmp = gdbCurrentDB.OpenRecordset(rName, dbOpenDynaset)
    sTmp = "Dynaset:"
  ElseIf gnRSType = gnRS_SNAPSHOT Then
    Set rsTmp = gdbCurrentDB.OpenRecordset(rName, dbOpenSnapshot)
    sTmp = "Snapshot:"
  ElseIf gnRSType = gnRS_PASSTHRU Then
    Set rsTmp = gdbCurrentDB.OpenRecordset(rName, dbOpenSnapshot, dbSQLPassThrough)
    sTmp = "Passthrough Snapshot:"
  End If
  Screen.MousePointer = vbDefault
  
  If gnFormType = gnFORM_NODATACTL Then
    If gnRSType = gnRS_TABLE Then
      Set frmTmp = New frmTableObj
      sTmp = "Tabla:"
    Else
      Set frmTmp = New frmDynaSnap
    End If
  ElseIf gnFormType = gnFORM_DATACTL Then
    Set frmTmp = New frmDataControl
  ElseIf gnFormType = gnFORM_DATAGRID Then
    Set frmTmp = New frmDataGrid
  End If
  
  Set frmTmp.mrsFormRecordset = rsTmp
  frmTmp.Caption = sTmp & rName
  frmTmp.Show

  MsgBar vbNullString, False
  
  Exit Sub
OpenTableErr:
  ShowError
End Sub

'------------------------------------------------------------
'abre una QueryDef con el tipo de formulario seleccionado
'por el usuario
'------------------------------------------------------------
Sub OpenQuery(rName As String, bTemp As Boolean)
  On Error GoTo OpenQueryErr

  Dim sTmp As String
  Dim rsTmp As Recordset
  Dim qdfTmp As QueryDef
  Dim sQueryType As String
  Dim frmTmp As Form
  Dim nDoIt As Integer
  Dim bReturnsRows As Boolean

  If bTemp Then
    Set qdfTmp = gdbCurrentDB.CreateQueryDef("", rName)
    If MsgBox(MSG17, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
      sTmp = InputBox(MSG18)
      If Len(sTmp) > 0 Then
        qdfTmp.Connect = sTmp
      End If
    End If
    'asume que no hay filas con las que empezar
    bReturnsRows = False
  Else
    Set qdfTmp = gdbCurrentDB.QueryDefs(rName)
    sQueryType = ActionQueryType(qdfTmp)
    If qdfTmp.Type <> dbQSQLPassThrough Then
      'no es una SQL específica así que tenemos que establecer ReturnsRecords
      If qdfTmp.Type = 0 Or qdfTmp.Type = dbQCrosstab Then
        bReturnsRows = True
      Else
        bReturnsRows = False
      End If
    Else
      'lo obtiene del QueryDef si es específica
      bReturnsRows = qdfTmp.ReturnsRecords
    End If
  End If
  
  If bReturnsRows And (gnRSType = gnRS_TABLE) Then
    Beep
    If MsgBox(MSG19, vbYesNo + vbQuestion) = vbYes Then
      frmMDI.tlbToolBar.Buttons("Dynaset").Value = tbrPressed         'restablece a Recordset
    Else
      Exit Sub
    End If
  End If
  
  
  If bReturnsRows Then
    SetQDFParams qdfTmp
MakeDynaset:
    Screen.MousePointer = vbHourglass
    If qdfTmp.Type = dbQSQLPassThrough Then
      MsgBar MSG16, True
      Set rsTmp = qdfTmp.OpenRecordset(dbOpenSnapshot, dbSQLPassThrough)
    ElseIf gnRSType = gnRS_SNAPSHOT Then
      MsgBar MSG20, True
      Set rsTmp = qdfTmp.OpenRecordset(dbOpenSnapshot)
    Else
      MsgBar MSG21, True
      Set rsTmp = qdfTmp.OpenRecordset(dbOpenDynaset)
    End If
    Screen.MousePointer = vbDefault
    
    If gnFormType = gnFORM_NODATACTL Then
      Set frmTmp = New frmDynaSnap
    ElseIf gnFormType = gnFORM_DATACTL Then
      Set frmTmp = New frmDataControl
      If qdfTmp.Parameters.Count > 0 Then
        frmTmp.mbIsParameterized = True
      End If
    ElseIf gnFormType = gnFORM_DATAGRID Then
      Set frmTmp = New frmDataGrid
    End If
    
    Set frmTmp.mrsFormRecordset = rsTmp
    If Len(qdfTmp.SQL) > 50 Then
      frmTmp.Caption = MSG22
    Else
      frmTmp.Caption = qdfTmp.SQL
    End If
    frmTmp.Show
    
  Else
    Screen.MousePointer = vbDefault
    If Len(sQueryType) > 0 Then
      nDoIt = MsgBox(MSG23 & sQueryType & MSG24, vbYesNo + vbQuestion)
    Else
      'sin nombre, de modo que intenta ejecutarlo
      nDoIt = vbYes
    End If
    If nDoIt = vbYes Then
      SetQDFParams qdfTmp
      Screen.MousePointer = vbHourglass
      MsgBar MSG25, True
      qdfTmp.Execute
      If gdbCurrentDB.RecordsAffected > 0 Then
        If gbTransPending Then gbDBChanged = True
      End If
    End If
  End If
  
  MsgBar vbNullString, False
  
  Exit Sub
OpenQueryErr:
  If Err = 3065 Or Err = 3078 Then
    'filas devueltas, luego intenta crear el Recordset
    Resume MakeDynaset
  End If
  ShowError
End Sub

'------------------------------------------------------------
'esta subrutina presenta todos los datos de los campos de la
'fila actual en los formularios table y dynasnap
'------------------------------------------------------------
Sub DisplayCurrentRecord(frm As Object, rec As Recordset, lCnt As Long, bNew As Integer)
  Dim i As Integer
  Dim sCurrStat As String
  Dim lCurrRec As Long
  Dim bNoInd As Integer

  On Error GoTo DCRErr

  Screen.MousePointer = vbHourglass

  sCurrStat = "Row "
   
  'comprueba si la tabla en uso tiene o no índices
  If rec.Type = dbOpenTable Then
    If gdbCurrentDB(rec.Name).Indexes.Count = 0 Then
      bNoInd = True
    End If
  End If
   
  'comprueba si el Recordset está vacío
  If rec.RecordCount > 0 Then
    lCurrRec = (lCnt * (rec.PercentPosition * 0.01)) + 1
  End If
     
  'comprueba el indicador BOF/EOF para saber si nos encontramos
  'en un registro válido
  If bNew Then
    If bNoInd Then
      sCurrStat = lCnt & " Filas"
    Else
      sCurrStat = lCurrRec & "/" & lCnt
    End If
  Else
    If rec.BOF Then
      sCurrStat = "(BOF)/" & lCnt
      ClearDataFields frm, rec.Fields.Count
    ElseIf rec.EOF Then
      sCurrStat = "(EOF)/" & lCnt
      ClearDataFields frm, rec.Fields.Count
    Else
      If bNoInd Then
        sCurrStat = lCnt & " Filas"
      Else
        sCurrStat = lCurrRec & "/" & lCnt
      End If
      'coloca los datos en los campos del formulario
      For i = 0 To rec.Fields.Count - 1
        If rec(i).Type = dbMemo Then
          If rec(i).FieldSize() < gnGETCHUNK_CUTOFF Then
            frm.txtFieldData(i).Text = StripNonAscii(vFieldVal(rec(i)))
          Else
            frm.txtFieldData(i).Text = StripNonAscii(vFieldVal(rec(i).GetChunk(0, gnGETCHUNK_CUTOFF)))
          End If
        ElseIf rec(i).Type = dbText Then
          frm.txtFieldData(i).Text = vFieldVal(rec(i))
        Else
          frm.txtFieldData(i).Text = vFieldVal(rec(i))
        End If
      Next
    End If
  End If
  If rec.Updatable = False Then sCurrStat = sCurrStat & MSG26
  frm.lblStatus.Caption = sCurrStat
  Screen.MousePointer = vbDefault
  Exit Sub

DCRErr:
  ShowError
  Resume Next    'podemos seguir y presentar tantos datos como sea posible
End Sub

'------------------------------------------------------------
'esta función comprueba si el nombre pasado existe en
'la colección Tabledefs o en la Querydefs,
'si se encuentra, pregunta si se elimina y devuelve False
'si el usuario selecciona eliminarlo o True si no,
'si no se encuentra, devuelve False
'------------------------------------------------------------
Function DupeTableName(rName As String) As Integer
  On Error GoTo DTNErr

  Dim tdf As TableDef
  Dim qdf As QueryDef
  Dim i As Integer

  For Each tdf In gdbCurrentDB.TableDefs
    If UCase(tdf.Name) = UCase(rName) Then
      If MsgBox(MSG27, vbYesNo + vbQuestion) = vbYes Then
        gdbCurrentDB.TableDefs.Delete rName
        DupeTableName = False
      Else
        DupeTableName = True
      End If
      Exit Function
    End If
  Next

  If gsDataType = gsMSACCESS Then
    For Each qdf In gdbCurrentDB.QueryDefs
      If UCase(qdf.Name) = UCase(rName) Then
        If MsgBox(MSG28, vbYesNo + vbQuestion) = vbYes Then
          gdbCurrentDB.QueryDefs.Delete rName
          DupeTableName = False
        Else
          DupeTableName = True
        End If
        Exit Function
      End If
    Next
  End If

  DupeTableName = False
  Exit Function

DTNErr:
  ShowError
  DupeTableName = False
End Function

'------------------------------------------------------------
'esta subrutina descarga todos los formularios excepto los
'formularios SQL, Tablas y MDI
'------------------------------------------------------------
Sub UnloadAllForms()
  On Error Resume Next
  
  Dim i As Integer
  
  'cierra todos los formularios excepto los formularios Tablas, SQL y MDI
  For i = Forms.Count - 1 To 1 Step -1
    Unload Forms(i)
  Next
End Sub

'------------------------------------------------------------
'esta subrutina recorre la colección de parámetros de las
'consultas con parámetros y pide al usuario un valor para
'cada uno de ellos
'------------------------------------------------------------
Sub SetQDFParams(rqdf As QueryDef)
  On Error GoTo SPErr
  
  Dim prm As Parameter
  Dim sTmp As String
  
  For Each prm In rqdf.Parameters
    'obtiene el valor del usuario
    sTmp = InputBox(MSG29, "'" & prm.Name & "':")
    'almacena el valor
    prm.Value = CVar(sTmp)
  Next
  
  Exit Sub
    
SPErr:
  ShowError
End Sub

'------------------------------------------------------------
'esta subrutina actualiza el formulario Error con los
'últimos errores
'------------------------------------------------------------
Sub RefreshErrors()
  On Error GoTo RErr
  
  Dim errObj As Error
  Dim i As Integer

  If DBEngine.Errors.Count = 0 Then
    MsgBox MSG30, 48
    Unload frmErrors
    Exit Sub
  End If

  frmErrors.Show
  frmErrors.lstErrors.Clear
  For i = 0 To DBEngine.Errors.Count - 1
    Set errObj = DBEngine.Errors(i)
    frmErrors.lstErrors.AddItem errObj.Number & vbTab & errObj.Source & vbTab & errObj.Description
  Next
  frmErrors.SetFocus

  Exit Sub
  
RErr:
  MsgBox MSG31, 48
  Unload frmErrors
  Exit Sub
End Sub

'------------------------------------------------------------
'esta subrutina agrega la base de datos recién abierta a
'la lista de las más recientes del menú Archivo
'------------------------------------------------------------
Sub AddMRU()
  On Error GoTo AMErr

  Dim i As Integer, j As Integer

  'primera investigación para ver si ya existe, y la cambia si existe
  For i = 1 To 8
    If UCase(Mid(frmMDI.mnuDBMRU(i).Caption, 4, Len(frmMDI.mnuDBMRU(i).Caption))) = UCase(gsDBName) Then
      For j = i To 2 Step -1
        frmMDI.mnuDBMRU(j).Caption = "&" & j & " " & Mid(frmMDI.mnuDBMRU(j - 1).Caption, 4, Len(frmMDI.mnuDBMRU(j - 1).Caption))
        frmMDI.mnuDBMRU(j).Tag = frmMDI.mnuDBMRU(j - 1).Tag
      Next
      GoTo Finish
    End If
  Next

  'no estaba, así que desplaza todos un lugar hacia abajo
  For i = 7 To 1 Step -1
    frmMDI.mnuDBMRU(i + 1).Caption = "&" & i + 1 & " " & Mid(frmMDI.mnuDBMRU(i).Caption, 4, Len(frmMDI.mnuDBMRU(i).Caption))
    frmMDI.mnuDBMRU(i + 1).Tag = frmMDI.mnuDBMRU(i).Tag
  Next

Finish:
  frmMDI.mnuDBMRU(1).Caption = "&1 " & gsDBName
  If Len(gdbCurrentDB.Connect) = 0 Then
    'gestiona el caso de Access, para el que no hay cadena de conexión
    frmMDI.mnuDBMRU(1).Tag = gsMSACCESS
  Else
    frmMDI.mnuDBMRU(1).Tag = gdbCurrentDB.Connect
  End If
  frmMDI.mnuBarMRU.Visible = True
  For i = 1 To 8
    If Len(Mid(frmMDI.mnuDBMRU(i).Caption, 4, Len(frmMDI.mnuDBMRU(i).Caption))) > 0 Then
      frmMDI.mnuDBMRU(i).Visible = True
    End If
  Next

  Exit Sub

AMErr:
  ShowError
End Sub

'------------------------------------------------------------
'esta subrutina separa las partes de una cadena de
'conexión ODBC y las asigna a las variables ODBC globales
'------------------------------------------------------------
Sub GetODBCConnectParts(rsConnect As String)
  On Error Resume Next
  
  Dim i As Integer
  Dim sTmp As String
  
  'procesa la cadena de conexión si los
  'valores vienen de los cuadros de diálogo ODBC
  If InStr(rsConnect, "=") Then
    i = 1
    While i <= Len(rsConnect) + 1
      If Mid(rsConnect, i, 1) = ";" Or i = Len(rsConnect) + 1 Then
        If Len(sTmp) > 0 And InStr(sTmp, "=") > 0 Then
          Select Case Mid(sTmp, 1, InStr(1, sTmp, "=") - 1)
            Case "DSN"
              gsODBCDatasource = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
            Case "DATABASE"
              gsODBCDatabase = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
            Case "DBQ"
              gsODBCDatabase = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
            Case "UID"
              gsODBCUserName = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
            Case "PWD"
              gsODBCPassword = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
            Case "Driver"
              gsODBCDriver = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
            Case "Server"
              gsODBCServer = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
             Case Else
              'nada
          End Select
        End If
        sTmp = vbNullString
      Else
        sTmp = sTmp + Mid(rsConnect, i, 1)
      End If
      i = i + 1
    Wend
  End If
End Sub

'------------------------------------------------------------
'esta es una subrutina genérica que agrega el nombre de
'cada elemento de una colección al control pasado
'------------------------------------------------------------
Sub ListItemNames(rcCollection As Object, rnCtl As Control, bClearList As Integer)
  On Error GoTo LINErr
  
  Dim objTmp As Object
  Dim i As Integer
  
  If bClearList Then
    rnCtl.Clear
  End If
  
  For Each objTmp In rcCollection
    rnCtl.AddItem objTmp.Name
  Next

  Exit Sub
  
LINErr:
  ShowError
End Sub

'------------------------------------------------------------
'esta subrutina cierra la base de datos actual y realiza la limpieza
'y restablecimiento de los controles, menús, etc.
'------------------------------------------------------------
Sub CloseCurrentDB()
  On Error GoTo DBCloseErr

  If gdbCurrentDB Is Nothing Then Exit Sub
    
  If gbDBChanged Then
    If MsgBox(MSG32, vbYesNo + vbQuestion) = vbYes Then
      gwsMainWS.CommitTrans
      gbDBChanged = False
    Else
      If MsgBox(MSG33, vbYesNo + vbQuestion) = vbYes Then
        gwsMainWS.Rollback
        gbDBChanged = False
      Else
        Beep
        MsgBox MSG34, 48
        Exit Sub
      End If
    End If
  End If

  frmMDI.Caption = "VisData"
  
  HideDBTools

  gbDBOpenFlag = False
  gbTransPending = False
  gsDBName = vbNullString
  gnReadOnly = False
  
  gdbCurrentDB.Close
  Set gdbCurrentDB = Nothing
  UnloadAllForms

  Exit Sub

DBCloseErr:
  ShowError
End Sub

'------------------------------------------------------------
'------------------------------------------------------------
Sub OpenLocalDB(bSilent As Boolean)
  On Error GoTo OpenError

  Dim sConnect As String
  Dim sDatabaseName As String
  Dim dbTemp As Database
  Dim sTmp As String

  sDatabaseName = gsDBName
  
  If Not bSilent Then
    Select Case gsDataType
      Case gsMSACCESS
        frmMDI.dlgCMD1.Filter = MSG49 & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG36
      Case gsDBASEIII, gsDBASEIV, gsDBASE5
        frmMDI.dlgCMD1.Filter = "Dbase (*.dbf)|*.dbf" & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG37
      Case gsFOXPRO20, gsFOXPRO25, gsFOXPRO26, gsFOXPRO30
        frmMDI.dlgCMD1.Filter = "FoxPro (*.dbf)|*.dbf" & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG38
      Case gsPARADOX3X, gsPARADOX4X, gsPARADOX5X
        frmMDI.dlgCMD1.Filter = "Paradox (*.db)|*.db" & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG39
      Case gsEXCEL50
        frmMDI.dlgCMD1.Filter = "Excel (*.xls)|*.xls" & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG40
      Case gsBTRIEVE
        frmMDI.dlgCMD1.Filter = "Btrieve (FILE.DDF)|FILE.DDF" & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG41
      Case gsTEXTFILES
        frmMDI.dlgCMD1.Filter = "Texto (*.txt)|*.txt" & MSG50
        frmMDI.dlgCMD1.DialogTitle = MSG42
    End Select

    frmMDI.dlgCMD1.FilterIndex = 1
    frmMDI.dlgCMD1.FileName = gsDBName  '""
    frmMDI.dlgCMD1.CancelError = True
    frmMDI.dlgCMD1.Flags = FileOpenConstants.cdlOFNFileMustExist
    frmMDI.dlgCMD1.ShowOpen

    If Len(frmMDI.dlgCMD1.FileName) > 0 Then
      gsDBName = frmMDI.dlgCMD1.FileName
    Else
      Exit Sub
    End If
  Else
    gsDBName = sDatabaseName
  End If
  
  If Len(gsDBName) = 0 Then
    MsgBar vbNullString, False
    Exit Sub
  End If

  MsgBar MSG43, True
  Screen.MousePointer = vbHourglass

  'establece la cadena de conexión
  If gsDataType = gsMSACCESS Then
    sConnect = vbNullString
  Else
    sConnect = gsDataType
  End If
  
  'establece el nombre de la base de datos para las bases de datos no
  ' Microsoft Access y Btrieve que vienen de CommonDialog
  If gsDataType <> gsMSACCESS And gsDataType <> gsBTRIEVE And _
     gsDataType <> gsEXCEL50 And (Not bSilent) Then
    'hay que eliminar el nombre del archivo en estas bases de datos
    sDatabaseName = StripFileName(gsDBName)
    gsDBName = sDatabaseName
  Else
    sDatabaseName = gsDBName
  End If

  GoTo OneMoreTry
  
GetPWD:
  Dim frmPWD As New frmDBPWD
  frmPWD.Show vbModal
  If Len(frmPWD.PWD) > 0 Then
    sConnect = ";pwd=" & frmPWD.PWD
    Unload frmPWD
    Set frmPWD = Nothing
    MsgBar MSG43, True
    Screen.MousePointer = vbHourglass
  Else
    'han cancelado el cuadro de diálogo de la contraseña, luego hay que salir
    Unload frmPWD
    Set frmPWD = Nothing
    Exit Sub
  End If

OneMoreTry:
  If (frmMDI.dlgCMD1.Flags And FileOpenConstants.cdlOFNReadOnly) = FileOpenConstants.cdlOFNReadOnly Then
    gnReadOnly = True
  Else
    gnReadOnly = False
  End If
  Set dbTemp = gwsMainWS.OpenDatabase(sDatabaseName, False, gnReadOnly, sConnect)
  If gbDBOpenFlag Then
    'guarda el nombre de la base de datos
    sTmp = gsDBName
    'lo recupera
    CloseCurrentDB
    gsDBName = sTmp
    If gbDBOpenFlag Then
      Beep
      MsgBox MSG35, 48
      Exit Sub
    End If
  End If

  'correcto
  frmMDI.Caption = "VisData:" & sDatabaseName
  Set gdbCurrentDB = dbTemp
  gbDBOpenFlag = True
  ShowDBTools
  RefreshTables Nothing
  gdbCurrentDB.QueryTimeout = glQueryTimeout

  AddMRU
  If gsDataType <> gsMSACCESS Then
    MsgBar MSG44, False
  End If
  Screen.MousePointer = vbDefault

  Exit Sub

AttemptRepair:
  Screen.MousePointer = vbHourglass
  MsgBar MSG45 & gsDBName, True
  DBEngine.RepairDatabase gsDBName
  Screen.MousePointer = vbDefault
  GoTo OneMoreTry

OpenError:
  Screen.MousePointer = vbDefault
  If Err = 3049 Then
    If MsgBox(Err.Description & vbCrLf & vbCrLf & MSG46, 4 + 48) = vbYes Then
      Resume AttemptRepair
    End If
  ElseIf Err = 3031 Then
    'base de datos protegida por contraseña
    Resume GetPWD
  End If
  gbDBOpenFlag = False
  gsDBName = vbNullString
  gsDataType = vbNullString
  gsODBCDatabase = vbNullString
  gsODBCUserName = vbNullString
  gsODBCPassword = vbNullString
  gsODBCDriver = vbNullString
  gsODBCServer = vbNullString
  If Err <> 32755 And Err <> 3049 Then   'comprueba si el CommonDialog ha sido cancelado
    ShowError
  End If
End Sub

'------------------------------------------------------------
'esta subrutina se utiliza para crear un nuevo directorio
'para uno de los tipos de datos ISAM locales
'------------------------------------------------------------
Sub NewLocalISAM()
   On Error GoTo NewISAMErr

   Dim sNewName As String
   Dim d As Database

GetNewDirName:
   sNewName = InputBox(MSG47, , sNewName)
   If Len(sNewName) = 0 Then Exit Sub

   If Mid(sNewName, Len(sNewName), 1) <> "\" Then sNewName = sNewName & "\"

   MkDir Mid(sNewName, 1, Len(sNewName) - 1)

   gsDBName = sNewName
   OpenLocalDB True

   If gbDBOpenFlag Then
     ShowDBTools
     RefreshTables Nothing
   End If

  Exit Sub

NewISAMErr:
  If Err = 75 Then Resume Next  'caso en el que el directorio ya existe
  If Err = 76 Then
    MsgBox MSG65, vbExclamation
    'vuelve a intentarlo
    Resume GetNewDirName
  End If
  ShowError
End Sub

'------------------------------------------------------------
'esta subrutina es llamada desde las opciones del menú
'de compactar del formulario MDI principal
'------------------------------------------------------------
Sub CompactDB(rnCompactVersion As Integer)
  On Error GoTo CompactAccErr

  Dim sOldName As String
  Dim sNewName As String
  Dim sNewName2 As String
  Dim nEncrypt As Integer

  'obtiene el nombre del archivo que se va a compactar
  frmMDI.dlgCMD1.Filter = MSG49
  frmMDI.dlgCMD1.DialogTitle = MSG48
  frmMDI.dlgCMD1.FilterIndex = 1
  frmMDI.dlgCMD1.Flags = FileOpenConstants.cdlOFNHideReadOnly
  frmMDI.dlgCMD1.ShowOpen
  If Len(frmMDI.dlgCMD1.FileName) > 0 Then
    sOldName = frmMDI.dlgCMD1.FileName
  Else
    Exit Sub
  End If

  'obtiene el nombre del archivo resultado de la compactación
  frmMDI.dlgCMD1.DialogTitle = MSG51
  frmMDI.dlgCMD1.FilterIndex = 1
  frmMDI.dlgCMD1.FileName = vbNullString
  frmMDI.dlgCMD1.CancelError = True
  frmMDI.dlgCMD1.Flags = FileOpenConstants.cdlOFNOverwritePrompt + FileOpenConstants.cdlOFNHideReadOnly
  frmMDI.dlgCMD1.ShowSave
  If Len(frmMDI.dlgCMD1.FileName) > 0 Then
    sNewName = frmMDI.dlgCMD1.FileName
    If Dir(sNewName) <> vbNullString And sOldName <> sNewName Then
      Kill sNewName
    End If
  Else
    Exit Sub
  End If

  If MsgBox(MSG52, vbYesNo + vbQuestion) = vbYes Then
    nEncrypt = dbEncrypt
  Else
    nEncrypt = dbDecrypt
  End If

  Screen.MousePointer = vbHourglass
  MsgBar MSG53 & sOldName & " -> " & sNewName, True
  'si no quieren sobrescribir el mismo archivo, tenemos que crear una nueva
  'MDB y cambiar su nombre después de que la compactación haya sido correcta
  If sOldName = sNewName Then
    sNewName2 = sNewName 'guarda el nombre nuevo
    sNewName = Left(sNewName, Len(sNewName) - 1) & "N"
  End If
  
  DBEngine.CompactDatabase sOldName, sNewName, dbLangGeneral, rnCompactVersion + nEncrypt
  
  'comprueba la sobrescritura de la mdb original
  If VBA.Right(sNewName, 1) = "N" Then
    Kill sNewName2             'acaba con la antigua
    Name sNewName As sNewName2 'cambia el nombre de la nueva por el de la original
    sNewName = sNewName2       'restablece el nombre correcto
  End If
  
  MsgBar vbNullString, False
  Screen.MousePointer = vbDefault

  If MsgBox(MSG54, vbYesNo + vbQuestion) = vbYes Then
    If gbDBOpenFlag Then
      CloseCurrentDB
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

CompactAccErr:
  If Err <> 32755 Then         'el usuario ha cancelado
    ShowError
  End If
End Sub

'------------------------------------------------------------
'esta subrutina hace algo de limpieza y termina VisData
'------------------------------------------------------------
Sub ShutDownVisData()
  On Error Resume Next

  Dim nRet As Integer

  'guarda toda la configuración de inicio actual
  SaveINISettings

  If gbDBChanged Then
    If MsgBox(MSG32, vbYesNo + vbQuestion) = vbYes Then
      gwsMainWS.CommitTrans
    End If
  End If

  UnloadAllForms
  gdbCurrentDB.Close
  'cierra el archivo de ayuda
  nRet = OSWinHelp(frmMDI.hwnd, App.HelpFile, HelpConstants.cdlHelpQuit, 0)
  
  End

End Sub
Sub NewMDB(rnVersion As Integer)
  On Error GoTo NewAccErr

  Dim sNewName As String
  Dim db As Database

  'obtiene el nombre del archivo resultado de la compactación
  frmMDI.dlgCMD1.DialogTitle = MSG55
  frmMDI.dlgCMD1.FilterIndex = 1
  frmMDI.dlgCMD1.Filter = MSG49
  frmMDI.dlgCMD1.FileName = vbNullString
  frmMDI.dlgCMD1.CancelError = True
  frmMDI.dlgCMD1.Flags = FileOpenConstants.cdlOFNOverwritePrompt + FileOpenConstants.cdlOFNHideReadOnly
  frmMDI.dlgCMD1.ShowSave
  If Len(frmMDI.dlgCMD1.FileName) > 0 Then
    sNewName = frmMDI.dlgCMD1.FileName
    If InStr(sNewName, ".") = 0 Then
      'agrega una extensión si el usuario no ha suministrado ninguna
      sNewName = sNewName & ".MDB"
    End If
    If Dir(sNewName) <> vbNullString Then
      Kill sNewName
    End If
  Else
    Exit Sub
  End If
  If Len(sNewName) = 0 Then Exit Sub

  Set db = CreateDatabase(sNewName, dbLangGeneral, rnVersion)
  db.Close

  gsDataType = gsMSACCESS
  gsDBName = sNewName
  OpenLocalDB True
  Exit Sub

NewAccErr:
  If Err <> 32755 Then         'el usuario ha cancelado
    ShowError
  End If
End Sub

Sub Export(rsFromTbl As String, rsToDB As String)

  On Error GoTo ExpErr

  Dim sConnect As String
  Dim sNewTblName As String
  Dim sDBName As String
  Dim nErrState As Integer
  Dim idxFrom As Index
  Dim idxTo As Index
  Dim sSQL As String              'copia local de la cadena SQL
  Dim sField As String
  Dim sFrom As String
  Dim sTmp As String
  Dim i As Integer

  If gnDataType = gnDT_SQLDB Then
    Set gExpDB = gwsMainWS.OpenDatabase(vbNullString, 0, 0, "odbc;")
    If gExpDB Is Nothing Then Exit Sub
  End If

  MsgBar MSG56 & "'" & rsFromTbl & "'", True

  nErrState = 1
  Select Case gnDataType
    Case gnDT_MSACCESS
      sConnect = "[;database=" & rsToDB & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(rsToDB)
    Case gnDT_PARADOX3X
      sDBName = StripFileName(rsToDB)
      sConnect = "[Paradox 3.X;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsPARADOX3X)
    Case gnDT_PARADOX4X
      sDBName = StripFileName(rsToDB)
      sConnect = "[Paradox 4.X;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsPARADOX4X)
    Case gnDT_FOXPRO26
      sDBName = StripFileName(rsToDB)
      sConnect = "[FoxPro 2.6;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsFOXPRO26)
    Case gnDT_FOXPRO25
      sDBName = StripFileName(rsToDB)
      sConnect = "[FoxPro 2.5;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsFOXPRO25)
    Case gnDT_FOXPRO20
      sDBName = StripFileName(rsToDB)
      sConnect = "[FoxPro 2.0;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsFOXPRO20)
    Case gnDT_DBASEIV
      sDBName = StripFileName(rsToDB)
      sConnect = "[dBase IV;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsDBASEIV)
    Case gnDT_DBASEIII
      sDBName = StripFileName(rsToDB)
      sConnect = "[dBase III;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsDBASEIII)
    Case gnDT_BTRIEVE
      sConnect = "[Btrieve;database=" & rsToDB & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(rsToDB, 0, 0, gsBTRIEVE)
    Case gnDT_EXCEL50, gnDT_EXCEL40, gnDT_EXCEL30
      sConnect = "[Excel 5.0;database=" & rsToDB & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(rsToDB, 0, 0, gsEXCEL50)
    Case gnDT_SQLDB
      sConnect = "[" & gExpDB.Connect & "]."
    Case gnDT_TEXTFILE
      sDBName = StripFileName(rsToDB)
      sConnect = "[Text;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsTEXTFILES)
  End Select
  If gnDataType = gnDT_MSACCESS Or gnDataType = gnDT_BTRIEVE Or _
     gnDataType = gnDT_SQLDB Or gnDataType = gnDT_EXCEL50 Or _
     gnDataType = gnDT_EXCEL40 Or gnDataType = gnDT_EXCEL30 Then
    With frmExpName
      .Label1.Caption = MSG57 & rsFromTbl & " ->"
      .Label2.Caption = MSG58 & rsToDB
      .txtTable.Text = rsFromTbl
    End With
    frmExpName.Show vbModal
      
    If Len(gExpTable) = 0 Then
      MsgBar vbNullString, False
      Exit Sub
    Else
      sNewTblName = gExpTable
    End If
  Else
    'obtiene la parte de la tabla desde el nombre del archivo
    'elimina la ruta
    For i = Len(rsToDB) To 1 Step -1
      If Mid(rsToDB, i, 1) = "\" Then
        Exit For
      End If
    Next
    sTmp = Mid(rsToDB, i + 1, Len(rsToDB))
    'elimina la extensión
    For i = 1 To Len(sTmp)
      If Mid(sTmp, i, 1) = "." Then
        Exit For
      End If
    Next
    sNewTblName = Left(sTmp, i - 1)
  End If
  Screen.MousePointer = vbHourglass
  If Len(rsFromTbl) > 0 Then
    gdbCurrentDB.Execute "select * into " & sConnect & StripOwner(sNewTblName) & " from " & StripOwner(rsFromTbl)

    If gnDataType <> gnDT_TEXTFILE Then
      nErrState = 2
      MsgBar MSG59 & " '" & sNewTblName & "'", True
      gExpDB.TableDefs.Refresh
      For Each idxFrom In gdbCurrentDB.TableDefs(rsFromTbl).Indexes
        Set idxTo = gExpDB.TableDefs(sNewTblName).CreateIndex(idxFrom.Name)
        With idxTo
          .Fields = idxFrom.Fields
          .Unique = idxFrom.Unique
          If gnDataType <> gnDT_SQLDB And gsDataType <> "ODBC" Then
            .Primary = idxFrom.Primary
          End If
        End With
        gExpDB.TableDefs(sNewTblName).Indexes.Append idxTo
      Next
    End If
    MsgBar vbNullString, False
    Screen.MousePointer = vbDefault
    MsgBox MSG60 & " '" & rsFromTbl & "'", 64
  Else
    sSQL = frmSQL.txtSQLStatement.Text
    sField = Mid(sSQL, 8, InStr(8, UCase(sSQL), "FROM") - 9)
    sFrom = " " & Mid(sSQL, InStr(UCase(sSQL), "FROM"), Len(sSQL))
    gdbCurrentDB.Execute "select " & sField & " into " & sConnect & sNewTblName & sFrom

    Screen.MousePointer = vbDefault
    MsgBar vbNullString, False
    MsgBox MSG61, 64
  End If

  Exit Sub

ExpErr:
  If Err = 3010 Then      'la tabla existe
    If MsgBox(MSG62, 32 + 1 + 256) = 1 Then
      gExpDB.TableDefs.Delete sNewTblName
      Resume
    Else
      Screen.MousePointer = vbDefault
      MsgBar vbNullString, False
      Exit Sub
    End If
  End If
 
  'elimina la nueva tabla si no se pueden crear los índices
  If nErrState = 2 Then
    gExpDB.TableDefs.Delete sNewTblName
  End If
  ShowError
End Sub

Sub Import(rsImpTblName As String)
  On Error GoTo ImpErr

  Dim sOldTblName As String, sNewTblName As String, sConnect As String
  Dim idxFrom As Index
  Dim idxTo As Index
  Dim nErrState As Integer
  Dim i As Integer

  sOldTblName = MakeTableName(rsImpTblName, False)
  sNewTblName = MakeTableName(rsImpTblName, True)

  Screen.MousePointer = vbHourglass
  MsgBar MSG63 & "'" & sNewTblName & "'", True

  nErrState = 1
  Select Case gnDataType
    Case gnDT_MSACCESS
      sConnect = "[;database=" & gImpDB.Name & "]."
    Case gnDT_PARADOX3X
      sConnect = "[Paradox 3.X;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsPARADOX3X)
    Case gnDT_PARADOX4X
      sConnect = "[Paradox 4.X;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsPARADOX4X)
    Case gnDT_FOXPRO26
      sConnect = "[FoxPro 2.6;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsFOXPRO26)
    Case gnDT_FOXPRO25
      sConnect = "[FoxPro 2.5;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsFOXPRO25)
    Case gnDT_FOXPRO20
      sConnect = "[FoxPro 2.0;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsFOXPRO20)
    Case gnDT_DBASEIV
      sConnect = "[dBase IV;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsDBASEIV)
    Case gnDT_DBASEIII
      sConnect = "[dBase III;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsDBASEIII)
    Case gnDT_BTRIEVE
      sConnect = "[Btrieve;database=" & gImpDB.Name & "]."
    Case gnDT_EXCEL50, gnDT_EXCEL40, gnDT_EXCEL30
      sConnect = "[Excel 5.0;database=" & gImpDB.Name & "]."
    Case gnDT_SQLDB
      sConnect = "[" & gImpDB.Connect & "]."
    Case gnDT_TEXTFILE
      sConnect = "[Text;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsTEXTFILES)
  End Select
  gdbCurrentDB.Execute "select * into " & sNewTblName & " from " & sConnect & sOldTblName

  If gnDataType <> gnDT_TEXTFILE And gnDataType <> gnDT_EXCEL50 And _
     gnDataType <> gnDT_EXCEL40 And gnDataType <> gnDT_EXCEL30 Then
    nErrState = 2
    MsgBar gdbCurrentDB.RecordsAffected & " filas importadas, creando índices para '" & sNewTblName & "'", True
    gdbCurrentDB.TableDefs.Refresh
    For Each idxFrom In gImpDB.TableDefs(sOldTblName).Indexes
      Set idxTo = gdbCurrentDB.TableDefs(sNewTblName).CreateIndex(idxFrom.Name)
      With idxTo
        .Fields = idxFrom.Fields
        .Unique = idxFrom.Unique
        If gnDataType <> gnDT_SQLDB And gsDataType <> gsSQLDB Then
          .Primary = idxFrom.Primary
        End If
      End With
      gdbCurrentDB.TableDefs(sNewTblName).Indexes.Append idxTo
    Next
  End If
    
  frmImpExp.lstTables.AddItem sNewTblName
'  frmTables.lstTables.AddItem sNewTblName
  Screen.MousePointer = vbDefault
  MsgBar vbNullString, False
  MsgBox MSG64 & "'" & sNewTblName & "'.", 64

  Exit Sub

NukeNewTbl:
  On Error Resume Next  'por si falla
  gdbCurrentDB.TableDefs.Delete sNewTblName
  ShowError
  Exit Sub
 
ImpErr:
  'elimina la nueva tabla si no se pueden crear los índices
  If nErrState = 2 Then
    Resume NukeNewTbl
  End If
  ShowError
End Sub

Function MakeTableName(fname As String, newname As Integer) As String
  On Error Resume Next
  Dim i As Integer, t As Integer
  Dim tmp As String

  If gnDataType = gnDT_SQLDB And newname Then
    i = InStr(1, fname, ".")
    If i > 0 Then
      tmp = Mid(fname, 1, i - 1) & "_" & Mid(fname, i + 1, Len(fname))
    End If
  ElseIf InStr(fname, "\") > 0 Then
    'elimina la ruta
    For i = Len(fname) To 1 Step -1
      If Mid(fname, i, 1) = "\" Then
        Exit For
      End If
    Next
    tmp = Mid(fname, i + 1, Len(fname))
    i = InStr(1, tmp, ".")
    If i > 0 Then
      tmp = Mid(tmp, 1, i - 1)
    End If
  Else
    tmp = fname
  End If

  If newname Then
    If DupeTableName(tmp) Then
      t = 1
      While DupeTableName(tmp + CStr(t))
        t = t + 1
      Wend
      tmp = tmp + CStr(t)
    End If
  End If

  MakeTableName = tmp

End Function


