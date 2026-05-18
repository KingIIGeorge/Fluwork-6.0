Attribute VB_Name = "modFriends"
Option Explicit
' Constantes para las funciones de depuración.
Global Const DEBUGTOKEN_DebugID = 1
Global Const DEBUGTOKEN_ClassName = 2

' Tipo definido por el usuario y utilizado por el
'   miembro Friend de la demostración.
Public Type udtDEMO
    intA As Integer
    lngB As Long
    strC As String
End Type

' timeGetTime es utilizado por la demostración de
' -----------     Implements y por el código de depuración
'   del ciclo de vida del objeto
'   de este módulo.
Declare Function timeGetTime Lib "winmm.dll" () As Long

' Almacenamiento de la colección global
'   para depurar los ciclos de vida de los objetos.
'   Se utiliza por los procedimientos DebugInit, DebugTerm,
'   y DebugShow (a continuación).
Private mcolDebug As New Collection

'         PROCEDIMIENTOS DE DEPURACIÓN
'
' DebugInit
' DebugTerm
' DebugShow
'
' Todos los objetos implementan la interfaz de IDebug
'   y la aceptan llamando a
' DebugInit(Me) en sus eventos
'   Initialize, y a DebugTerm Me
'   en sus eventos Terminate. Puede utilizar
'   DebugShow en la ventana Inmediato
'   para enumerar un objeto activo, todos los objetos
'   activos, o todos los objetos de una clase.
'
' Todos los objetos de este proyecto están
'   configurados para utilizar dichas funciones.
'
' -------------------------------------
' Todos los objetos llaman a DebugInit,
'   en su evento Initialize.
'   DebugInit agrega una cadena de depuración para el
'   objeto a la colección global y
'   devuelve un DebugID único para el
'   objeto. El método muestra opcionalmente
'   la cadena de depuración en la ventana Inmediato
'   (el valor predeterminado es True).
'
Public Function DebugInit(ByVal obj As Object, _
        Optional ByVal ShowImmediate As Boolean = True) As Long
    Dim lngDebugID As Long
    Dim strDebug As String
    
    ' Obtiene un número ID único.
    lngDebugID = GetDebugID
    ' La cadena de depuración de cada objeto
    '   muestra el DebugID, el nombre
    '   de la clase del objeto y la hora en que
    '   fue iniciado
    '   (número de segundos desde que se
    '   creó el primer objeto de depuración,
    '   expresados como Double, con los
    '   milisegundos como parte
    '   fraccional).
    strDebug = lngDebugID & " " _
        & TypeName(obj) _
        & " (creado a las " & DebugTime & ")"
    '
    ' Agrega la cadena a la colección,
    '   utilizando el ID único como clave.
    mcolDebug.Add strDebug, CStr(lngDebugID)
    '
    ' La acción predeterminada es mostrar la cadena de
    '   depuración en la ventana Inmediato.
    If ShowImmediate Then Debug.Print strDebug
    '
    ' Devolución de DebugID. El objeto tiene
    '   que guardar esto como parte de la
    '   implementación de IDebug.
    DebugInit = lngDebugID
End Function

' Se llama a DebugTerm por cada objeto,
'   en su evento Terminate.
'   DebugTerm elimina la cadena de depuración
'   del objeto de la colección global,
'   y opcionalmente (lo es de forma predeterminada)
'   muestra la cadena de depuración en la
'   ventana Inmediato.
'
Public Sub DebugTerm(ByVal obj As Object, _
        Optional ByVal ShowImmediate As Boolean = True)
    
    Dim idbg As IDebug
    
    On Error Resume Next
    '
    ' Obtiene una referencia a la interfaz de
    '   IDebug del objeto.
    Set idbg = obj
    If Err.Number <> 0 Then
        MsgBox TypeName(obj) & " no implementa IDebug; imposible guardar la terminación.", , "DebugTerm"
        Exit Sub
    End If
    '
    ' La acción predeterminada es mostrar la cadena
    '   de depuración en la ventana Inmediato.
    If ShowImmediate Then Debug.Print _
        mcolDebug(CStr(idbg.DebugID)) _
        & " (Terminado a las " & DebugTime & ")"
    '
    ' Elimina la cadena de la
    '   colección.
    mcolDebug.Remove CStr(idbg.DebugID)
End Sub

' DebugShow muestra las cadenas de depuración
' ---------     de toda la lista
'   de objetos activos, de todos los objetos
'   activos de una clase, o de un objeto concreto.
'   Llame a DebugShow desde la ventana
'   Inmediato sin argumentos (todos),
'   con un nombre de clase (todos los de una clase,
'   una referencia de objeto (dicho objeto),
'   o el DebugID de un objeto
'   (dicho objeto).
'
Public Sub DebugShow(Optional ByVal What As Variant)
    Dim vnt As Variant
    Dim idbg As IDebug
    
    On Error GoTo NoShow
    ' Si no se pasan argumentos, presenta
    '   todos los objetos activos. (Sería útil
    '   tener un segundo parámetro
    '   Filename que permitiera
    '   volcar esto en un archivo; o quizás
    '   en el Portapapeles).
    If IsMissing(What) Then
        What = "<All>"
        For Each vnt In mcolDebug
            Debug.Print vnt
        Next
    '
    ' Si se aporta un objeto, utiliza su
    '   DebugID para buscar su cadena de
    '   depuración.
    ElseIf IsObject(What) Then
        On Error Resume Next
        '
        ' Obtiene una referencia a la interfaz
        '   IDebug del objeto.
        Set idbg = What
        If Err.Number <> 0 Then
            MsgBox TypeName(What) & " no implementa IDebug; imposible guardar el registro de depuración.", , "DebugShow"
            Exit Sub
        End If
       
        Debug.Print mcolDebug(CStr(idbg.DebugID))
    '
    ' Si se suministra un número, asume que es
    '   un DebugID y lo utiliza para buscar la
    '   cadena.
    ElseIf IsNumeric(What) Then
        Debug.Print mcolDebug(CStr(What))
    '
    ' Si no es un número, asume que es
    '   una cadena que contiene el nombre de la
    '   clase y presenta todos los objetos de
    '   dicha clase.
    Else
        For Each vnt In mcolDebug
            If What = GetDebugToken(vnt, DEBUGTOKEN_ClassName) Then
                Debug.Print vnt
            End If
        Next
    End If
    Exit Sub
    
NoShow:
    If IsObject(What) Then
        MsgBox "Imposible mostrar información. ¿Permite depuración este objeto?", , "DebugShow"
    Else
        MsgBox "Imposible mostrar información de " _
            & What & ". ¿Permite depuración este objeto?", , "DebugShow"
    End If
End Sub

' GetDebugString devuelve la cadena de un
'  objeto desde la
'  colección global.
'
Public Function GetDebugString(ByVal obj As Object) As String
    Dim idbg As IDebug
    
    On Error Resume Next
    '
    ' Obtiene una referencia a la interfaz
    '   IDebug del objeto.
    Set idbg = obj
    GetDebugString = mcolDebug(CStr(idbg.DebugID))
End Function

' GetDebugID se utiliza para asignar a cada objeto
'   un número de ID único, por
'   propósitos de depuración.
Public Function GetDebugID() As Long
    Static lngLastID As Long
    lngLastID = lngLastID + 1
    GetDebugID = lngLastID
End Function

' GetDebugToken extrae la cadena de depuración
'   de un objeto y
'   devuelve dicha cadena. Las partes
'   están separadas por un espacio.
'   (1) DebugID
'   (2) nombre de clase
'
' Hay más partes, pero no
'   nos interesan.
'
Public Function GetDebugToken( _
        ByVal DebugString As String, _
        ByVal TokenNumber As Integer) As String

    Dim inx1 As Long
    Dim inx2 As Long
    Dim ct As Integer
    
    If TokenNumber <= 0 Then
        Err.Raise vbObjectError + 1060, , _
            "Numero incorrecto de elemento en GetDebugToken"
    Else
        inx2 = 1
        For ct = 1 To TokenNumber
            inx1 = inx2
            inx2 = InStr(inx1, DebugString, " ")
            If inx2 = 0 Then Exit For
        Next
        If inx2 = 0 Then
            GetDebugToken = ""
        Else
            GetDebugToken = Mid$(DebugString, inx1 + 1, inx2 - inx1)
        End If
    End If
End Function
        
' DebugTime utiliza la API timeGetTime para
'   obtener los milisegundos desde
'   que se inició el equipo.  Esto se convierte
'   en un Double que contiene el
'   número de segundos desde que se creó
'   el primer objeto de depuración (s.mmm),
'   a lo obtenido la primera vez que se utiliza esta
'   función se le llama hora base. (Esto hace que
'   estos valores sean más útiles que el número
'   de milisegundos desde el inicio,
'   (1) tiende a ser un
'   número muy alto y (2) puede
'   ser negativo, como se explica a continuación.)
'
Public Function DebugTime() As Double
    Static timeBase As Double
    Dim timeCurrent As Double
    
    If timeBase = 0 Then
        ' Inicializa la hora base.  (El bucle
        '   permite que el tiempo devuelto
        '   por timeGetTime
        '   puede pasar otra vez por cero, si el
        '   equipo siguiese en funcionamiento el
        '   tiempo necesario).
        Do While timeBase = 0
            timeBase = timeGetTime
        Loop
        '
        ' El valor devuelto por timeGetTime
        '   puede ser negativo (vea la siguiente
        '   nota) si el equipo lleva mucho tiempo
        '   en funcionamiento.
        '   Esto corrige esta situación.
        If timeBase < 0 Then
            timeBase = timeBase + 4294967296#
        End If
    End If
    '
    timeCurrent = timeGetTime
    '
    ' Corrige el valor negativo, si
    '   fuera necesario.
    If timeCurrent < 0 Then
        timeCurrent = timeCurrent + 4294967296#
    End If
    '
    ' Controla el caso en el que timeGetTime
    '   pasa por cero.
    If timeCurrent < timeBase Then
        DebugTime = (timeCurrent + 4294967296# - timeBase) / 1000#
    Else
        DebugTime = (timeCurrent - timeBase) / 1000#
    End If
End Function
' ----------- timeGetTime -----------
' El número de milisegundos desde el
'   último inicio es un entero de 4 bytes
'   sin signo, lo que significa que puede ser
'   mayor de lo que cabe en un Long.
'   Cuando se pasa del mayor número positivo
'   que cabe en un Long, 2147483647,
'   VB piensa que el número se ha desbordado
'   y se ha vuelto negativo.
'   Una vez negativo, sigue
'   creciendo desde el
'   número negativo mayor que cabe en un
'   Long hasta cero, y después vuelve a
'   ser un número positivo.
'
' Esto crea un figura en 'diente de sierra',
'   y funciona bien para diferencias de
'   tiempo(que es lo que
'   DebugTime calcula), excepto
'   cuando se produce el paso
'   del desbordamiento.
'
' DebugTime resuelve este problema
'   poniendo el número en un contenedor
'   más grande, un Double.  Si el número
'   es negativo, puede ser convertido
'   en el número que tendría que haber sido
'   agregándole 4294967296.
