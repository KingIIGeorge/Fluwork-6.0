Attribute VB_Name = "Module1"
'*** Módulo global de la aplicación de ejemplo Bloc de notas MDI.  ***
'*********************************************************************
Option Explicit

' Tipo definido por el usuario para almacenar información sobre los formularios hijos
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Public FState()  As FormState           ' Matriz de tipos definidos por el usuario
Public Document() As New frmNotePad     ' Matriz de objetos formularios hijo
Public gFindString As String            ' Almacena el texto de búsqueda.
Public gFindCase As Integer             ' Búsqueda sensible a mayúsculas y minúsculas
Public gFindDirection As Integer        ' Dirección de búsqueda.
Public gCurPos As Integer               ' Almacena la posición del cursor.
Public gFirstTime As Integer            ' Posición inicial.
Public gToolsHidden As Boolean          ' Estado de la barra de herramientas.
Public Const ThisApp = "MDINote"        ' Constante del Registro del sistema.
Public Const ThisKey = "Recent Files"   ' Constante del Registro del sistema.

Function AnyPadsLeft() As Integer
    Dim i As Integer        ' Variable contador

    ' Recorre la matriz de documentos.
    ' Vuelve aquí si al menos hay un documento abierto.
    For i = 1 To UBound(Document)
        If Not FState(i).Deleted Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next
End Function

Sub EditCopyProc()
    ' Copia el texto seleccionado en el Portapapeles.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
End Sub

Sub EditCutProc()
    ' Copia el texto seleccionado en el Portapapeles.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Elimina el texto seleccionado.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub

Sub EditPasteProc()
    ' Coloca el texto del Portapapeles en el control activo.
    frmMDI.ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub

Sub FileNew()
    Dim fIndex As Integer

    ' Busca el siguiente índice disponible y muestra el formulario hijo.
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = "Sinnombr:" & fIndex
    Document(fIndex).Show

    ' Asegura que los botones de edición de la barra de herramientas son visibles.
    frmMDI.imgCutButton.Visible = True
    frmMDI.imgCopyButton.Visible = True
    frmMDI.imgPasteButton.Visible = True
End Sub

Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Recorre la matriz de documentos. Si alguno de los
    ' documentos ha sido eliminado, devuelve dicho índice.
    For i = 1 To ArrayCount
        If FState(i).Deleted Then
            FindFreeIndex = i
            FState(i).Deleted = False
            Exit Function
        End If
    Next

    ' Si ninguno de los elementos de la matriz de documentos ha sido
    ' eliminado, incrementa en uno el rango de la matriz de documentos y de estados
    ' y devuelve el índice del nuevo elemento.
    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve FState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

Sub FindIt()
    Dim intStart As Integer
    Dim intPos As Integer
    Dim strFindString As String
    Dim strSourceString As String
    Dim strMsg As String
    Dim intResponse As Integer
    Dim intOffset As Integer
    
    ' Establece la variable de desplazamiento basándose en la posición del cursor.
    If (gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    ' Lee la variable pública de la posición de inicio.
    If gFirstTime Then intOffset = 0
    ' Asigna el valor de inicio.
    intStart = frmMDI.ActiveForm.ActiveControl.SelStart + intOffset
        
    ' Si no distingue entre mayúsculas y minúsculas, convierte la cadena en mayúsculas
    If gFindCase Then
        strFindString = gFindString
        strSourceString = frmMDI.ActiveForm.ActiveControl.Text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(frmMDI.ActiveForm.ActiveControl.Text)
    End If
            
    ' Busca la cadena.
    If gFindDirection = 1 Then
        intPos = InStr(intStart + 1, strSourceString, strFindString)
    Else
        For intPos = intStart - 1 To 0 Step -1
            If intPos = 0 Then Exit For
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then Exit For
        Next
    End If

    ' Si se encuentra la cadena...
    If intPos Then
        frmMDI.ActiveForm.ActiveControl.SelStart = intPos - 1
        frmMDI.ActiveForm.ActiveControl.SelLength = Len(strFindString)
    Else
        strMsg = "Imposible buscar " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(strMsg, 0, App.Title)
    End If
    
    ' Restablece las variables públicas
    gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart
    gFirstTime = False
End Sub

Sub GetRecentFiles()
    ' Este procedimiento demuestra la utilización de la función GetAllSettings,
    ' que devuelve una matriz de valores desde el Registro del sistema. En este caso,
    ' el Registro del sistema contiene los archivos más recientemente abiertos. Utilice la
    ' instrucción SaveSetting para escribir los nombres de los archivos más recientes.
    ' Dicha instrucción se utiliza en el procedimiento WriteRecentFiles.
    Dim i, j As Integer
    Dim varFiles As Variant ' Variable para almacenar la matriz devuelta.
    
    ' Obtiene los archivos más recientes desde el Registro del sistema
    ' mediante la instrucción GetAllSettings.
    ' ThisApp y ThisKey son constantes definidas en este módulo.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For i = 0 To UBound(varFiles, 1)
        
        frmMDI.mnuRecentFile(0).Visible = True
        frmMDI.mnuRecentFile(i).Caption = varFiles(i, 1)
        frmMDI.mnuRecentFile(i).Visible = True
            ' Recorre todos los documentos y actualiza cada menú.
            For j = 1 To UBound(Document)
                If Not FState(j).Deleted Then
                    Document(j).mnuRecentFile(0).Visible = True
                    Document(j).mnuRecentFile(i + 1).Caption = varFiles(i, 1)
                    Document(j).mnuRecentFile(i + 1).Visible = True
                End If
            Next j
    Next i

End Sub

Sub OptionsToolbarProc(CurrentForm As Form)
    ' Alterna la comprobación
    CurrentForm.mnuOptionsToolbar.Checked = Not CurrentForm.mnuOptionsToolbar.Checked
    ' Si no es el formulario MDI, establece la comprobación del formulario MDI.
    If Not TypeOf CurrentForm Is MDIForm Then
        frmMDI.mnuOptionsToolbar.Checked = CurrentForm.mnuOptionsToolbar.Checked
    End If
    ' Alterna la visibilidad de la barra de herramientas según este valor.
    If CurrentForm.mnuOptionsToolbar.Checked Then
        frmMDI.picToolbar.Visible = True
    Else
        frmMDI.picToolbar.Visible = False
    End If
End Sub

Sub WriteRecentFiles(OpenFileName)
    ' Este procedimiento utiliza la instrucción SaveSettings para escribir los nombres
    ' de los archivos más recientes en el Registro del sistema. La instrucción SaveSetting
    ' requiere tres parámetros. Dos de ellos son almacenados como
    ' constantes y se definen en este módulo.  La función GetAllSettings
    ' se utiliza en el procedimiento GetRecentFiles para obtener los
    ' nombres de archivo almacenados en este procedimiento.
    
    Dim i, j As Integer
    Dim strFile, key As String

    ' Copia RecentFile1 en RecentFile2, y así sucesivamente.
    For i = 3 To 1 Step -1
        key = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, key)
        If strFile <> "" Then
            key = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, key, strFile
        End If
    Next i
  
    ' Escribe al archivo abierto como primer archivo más reciente.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

