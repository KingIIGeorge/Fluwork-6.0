Attribute VB_Name = "Module2"
'*** Módulo estándar que contiene procedimientos para trabajar con archivos.   ***
'*** Forma parte de la aplicación de ejemplo Bloc de notas MDI.                ***
'*********************************************************************************
Option Explicit

Sub FileOpenProc()
    Dim intRetVal
    On Error Resume Next
    Dim strOpenFileName As String
    frmMDI.CMDialog1.Filename = ""
    frmMDI.CMDialog1.ShowOpen
    If Err <> 32755 Then    ' El usuario eligió Cancelar.
        strOpenFileName = frmMDI.CMDialog1.Filename
        ' Si el archivo es mayor de 65K, no puede
        ' ser abierto, de modo que cancela la operación.
        If FileLen(strOpenFileName) > 65000 Then
            MsgBox "El archivo es muy grande para abrirlo."
            Exit Sub
        End If
        
        OpenFile (strOpenFileName)
        UpdateFileMenu (strOpenFileName)
        ' Muestra las barras de herramientas si no son visibles.
        If gToolsHidden Then
            frmMDI.imgCutButton.Visible = True
            frmMDI.imgCopyButton.Visible = True
            frmMDI.imgPasteButton.Visible = True
            gToolsHidden = False
        End If
    End If
End Sub

Function GetFileName(Filename As Variant)
    ' Presenta un cuadro de diálogo Guardar como y devuelve un nombre de archivo.
    ' Si el usuario elige Cancelar, devuelve una cadena vacía.
    On Error Resume Next
    frmMDI.CMDialog1.Filename = Filename
    frmMDI.CMDialog1.ShowSave
    If Err <> 32755 Then    ' El usuario eligió Cancelar.
        GetFileName = frmMDI.CMDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function

Function OnRecentFilesList(Filename) As Integer
  Dim i     ' Variable contador.

  For i = 1 To 4
    If frmMDI.mnuRecentFile(i).Caption = Filename Then
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
    
    ' Modifica el título del formulario y lo presenta.
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = UCase(Filename)
    Document(fIndex).Text1.Text = Input(LOF(1), 1)
    FState(fIndex).Dirty = False
    Document(fIndex).Show
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
    strContents = frmMDI.ActiveForm.Text1.Text
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
        frmMDI.ActiveForm.Caption = UCase(Filename)
        ' Restablece el indicador Dirty.
        FState(frmMDI.ActiveForm.Tag).Dirty = False
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

