Attribute VB_Name = "Module1"
'*** Módulo global de la aplicación de ejemplo Bloc de notas SDI.  ***
'*********************************************************************
Option Explicit

' Tipo definido por el usuario para almacenar información sobre los formularios hijos
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Public FState As FormState              ' Matriz de tipos definidos por el usuario
Public gFindString As String            ' Almacena el texto de búsqueda.
Public gFindCase As Integer             ' Búsqueda sensible a mayúsculas y minúsculas
Public gFindDirection As Integer        ' Dirección de búsqueda.
Public gCurPos As Integer               ' Almacena la posición del cursor.
Public gFirstTime As Integer            ' Posición inicial.
Public Const ThisApp = "MDINote"        ' Constante del Registro del sistema.
Public Const ThisKey = "Recent Files"   ' Constante del Registro del sistema.

Sub EditCopyProc()
    ' Copia el texto seleccionado en el Portapapeles.
    Clipboard.SetText frmSDI.txtNote.SelText
End Sub

Sub EditCutProc()
    ' Copia el texto seleccionado en el Portapapeles.
    Clipboard.SetText frmSDI.txtNote.SelText
    ' Elimina el texto seleccionado.
    frmSDI.txtNote.SelText = ""
End Sub

Sub EditPasteProc()
    ' Coloca el texto del Portapapeles en el control activo.
    frmSDI.txtNote.SelText = Clipboard.GetText()
End Sub

Sub FileNew()
    Dim intResponse As Integer
    
    ' Si el archivo ha sido modificado, lo guarda
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    ' Borra el cuadro de texto y actualiza el título.
    frmSDI.txtNote.Text = ""
    frmSDI.Caption = "Bloc de notas SDI - Sintítulo"
End Sub
Function FileSave() As Integer
    Dim strFilename As String

    If frmSDI.Caption = "Bloc de notas SDI - Sintítulo" Then
        ' El archivo no ha sido guardado aún.
        ' Obtiene el nombre del archivo, y después llama al procedimiento de guardar.
        strFilename = GetFileName(strFilename)
    Else
        ' El título del formulario contiene el nombre del archivo abierto.
        strFilename = Right(frmSDI.Caption, Len(frmSDI.Caption) - 20)
    End If
    ' Llama al procedimiento de guardar. Si Filename está vacía, es porque
    ' el usuario eligió Cancelar en el cuadro de diálogo Guardar como; si no,
    ' guarda el archivo.
    If strFilename <> "" Then
        SaveFileAs strFilename
        FileSave = True
    Else
        FileSave = False
    End If
End Function


Sub FindIt()
    Dim intStart As Integer
    Dim intPos As Integer
    Dim strFindString As String
    Dim strSourceString As String
    Dim strMsg As String
    Dim intResponse As Integer
    Dim intOffset As Integer
    
    ' Establece la variable de desplazamiento de acuerdo con la posición del cursor.
    If (gCurPos = frmSDI.txtNote.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    ' Lee la variable pública de la posición de inicio.
    If gFirstTime Then intOffset = 0
    ' Asigna el valor de inicio.
    intStart = frmSDI.txtNote.SelStart + intOffset
        
    ' Si no distingue entre mayúsculas y minúsculas, convierte la cadena en mayúsculas
    If gFindCase Then
        strFindString = gFindString
        strSourceString = frmSDI.txtNote.Text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(frmSDI.txtNote.Text)
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
        frmSDI.txtNote.SelStart = intPos - 1
        frmSDI.txtNote.SelLength = Len(strFindString)
    Else
        strMsg = "Imposible encontrar " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(strMsg, 0, App.Title)
    End If
    
    ' Restablece las variables públicas
    gCurPos = frmSDI.txtNote.SelStart
    gFirstTime = False
End Sub

Sub GetRecentFiles()
    ' Este procedimiento demuestra la utilización de la función GetAllSettings,
    ' que devuelve una matriz de valores desde el Registro del sistema. En este caso,
    ' el Registro del sistema contiene los archivos más recientemente abiertos.  Utilice la
    ' instrucción SaveSetting para escribir los nombres de los archivos más recientes.
    ' Dicha instrucción se utiliza en el procedimiento WriteRecentFiles.
    Dim i As Integer
    Dim varFiles As Variant ' Variable para almacenar la matriz devuelta.
    
    ' Obtiene los archivos más recientes desde el Registro del sistema
    ' mediante la instrucción GetAllSettings.
    ' ThisApp y ThisKey son constantes definidas en este módulo.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For i = 0 To UBound(varFiles, 1)
        frmSDI.mnuRecentFile(0).Visible = True
        frmSDI.mnuRecentFile(i + 1).Caption = varFiles(i, 1)
        frmSDI.mnuRecentFile(i + 1).Visible = True
    Next i
End Sub
Sub ResizeNote()
    ' Expande el cuadro de texto para que ocupe toda el área interna del formulario.
    If frmSDI.picToolbar.Visible Then
        frmSDI.txtNote.Height = frmSDI.ScaleHeight - frmSDI.picToolbar.Height
        frmSDI.txtNote.Width = frmSDI.ScaleWidth
        frmSDI.txtNote.Top = frmSDI.picToolbar.Height
    Else
        frmSDI.txtNote.Height = frmSDI.ScaleHeight
        frmSDI.txtNote.Width = frmSDI.ScaleWidth
        frmSDI.txtNote.Top = 0
    End If
End Sub


Sub WriteRecentFiles(OpenFileName)
    ' Este procedimiento utiliza la instrucción SaveSettings para escribir los nombres
    ' de los archivos más recientes en el Registro del sistema. La instrucción SaveSetting
    ' requiere tres parámetros. Dos de ellos son almacenados como
    ' constantes y se definen en este módulo. La función GetAllSettings
    ' se utiliza en el procedimiento GetRecentFiles para obtener los
    ' nombres de archivo almacenados en este procedimiento.
    
    Dim i As Integer
    Dim strFile As String
    Dim strKey As String

    ' Copia RecentFile1 en RecentFile2, y así sucesivamente.
    For i = 3 To 1 Step -1
        strKey = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, strKey)
        If strFile <> "" Then
            strKey = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, strKey, strFile
        End If
    Next i
  
    ' Escribe al archivo abierto como primer archivo más reciente.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

