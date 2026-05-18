Attribute VB_Name = "Module2"
'*** Módulo estándar que contiene procedimientos para trabajar con archivos.   ***
'*** Forma parte de la aplicación de muestra Bloc de notas SDI.                ***
'*********************************************************************************
Option Explicit

Sub FileOpenProc()
    Dim intRetVal
    Dim intResponse As Integer
    Dim strOpenFileName As String
    
    ' Si el archivo ha sido modificado, lo guarda
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    On Error Resume Next
    
    frmSDI.CMDialog1.Filename = ""
    frmSDI.CMDialog1.Filter = "Texto (*.txt)|*.txt|Todos los archivos  (*.*)|*.*|"
    frmSDI.CMDialog1.ShowOpen
    If Err <> 32755 Then    ' El usuario eligió Cancelar.
        strOpenFileName = frmSDI.CMDialog1.Filename
        ' Si el archivo es mayor de 65K, no puede ser
        ' abierto, de modo que cancela la operación.
        If FileLen(strOpenFileName) > 65000 Then
            MsgBox "El archivo es muy grande para abrir."
            Exit Sub
        End If
        
        OpenFile (strOpenFileName)
        UpdateFileMenu (strOpenFileName)
    End If
End Sub

Function GetFileName(Filename As Variant)
    ' Presenta un cuadro de diálogo Guardar como y devuelve un nombre de archivo.
    ' Si el usuario elige Cancelar, devuelve una cadena vacía.
    On Error Resume Next
    frmSDI.CMDialog1.Filename = Filename
    frmSDI.CMDialog1.Filter = "Texto (*.txt)|*.txt|Todos los archivos  (*.*)|*.*|"
    frmSDI.CMDialog1.ShowSave
    If Err <> 32755 Then    ' El usuario eligió Cancelar.
        GetFileName = frmSDI.CMDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function

Function OnRecentFilesList(Filename) As Integer
    Dim i         ' Variable contador.

    For i = 1 To 4
        If frmSDI.mnuRecentFile(i).Caption = Filename Then
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
    OnRecentFilesList = False
End Function

Sub OpenFile(Filename)
    Dim fIndex As Integer
    
    On Error Resume Next
    ' Abre el archivo seleccionado.
    Open Filename For Input As #1
    If Err Then
        MsgBox "Imposible abrir el archivo: " + Filename
        Exit Sub
    End If
    ' Cambia el puntero del mouse a reloj de arena.
    Screen.MousePointer = 11
    
    ' Modifica el título del formulario y presenta el nuevo texto.
    frmSDI.Caption = "Bloc de notas SDI - " & UCase(Filename)
    frmSDI.txtNote.Text = Input(LOF(1), 1)
    FState.Dirty = False
    Close #1
    ' Restablece el puntero del mouse.
    Screen.MousePointer = 0
End Sub

Sub SaveFileAs(Filename)
    On Error Resume Next
    Dim strContents As String

    ' Abre el archivo.
    Open Filename For Output As #1
    ' Coloca el contenido del bloc de notas en una variable.
    strContents = frmSDI.txtNote.Text
    ' Presenta el puntero reloj de arena.
    Screen.MousePointer = 11
    ' Escribe el contenido de la variable en un archivo.
    Print #1, strContents
    Close #1
    ' Restablece el puntero del mouse.
    Screen.MousePointer = 0
    ' Establece el título del formulario.
    If Err Then
        MsgBox Error, 48, App.Title
    Else
        frmSDI.Caption = "Bloc de notas SDI - " & Filename
        ' Restablece el indicador Dirty.
        FState.Dirty = False
    End If
End Sub

Sub UpdateFileMenu(Filename)
        Dim intRetVal As Integer
        ' Comprueba si el archivo abierto se encuentra en la matriz de controles del menú Archivo.
        intRetVal = OnRecentFilesList(Filename)
        If Not intRetVal Then
            ' Escribe el nombre del archivo abierto en el Registro del sistema.
            WriteRecentFiles (Filename)
        End If
        ' Actualiza la lista de los archivos abiertos recientemente en la matriz de controles
        ' del menú Archivo.
        GetRecentFiles
End Sub

