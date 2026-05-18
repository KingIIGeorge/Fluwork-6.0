Attribute VB_Name = "ModOLECont"
Option Explicit

Public MDINew As Integer

Sub NewObject()
    MDINew = True
    NewOleForm
    If ParentForm.ActiveForm.OLE1.OLEType = vbOLENone Then
        Unload ParentForm.ActiveForm
    End If
End Sub

Sub DisplayInstructions()
    ' Declara variables locales.
    Dim MsgText
    Dim PB
    ' Inicializa la variable de ruptura de párrafo.
    PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
    ' Presenta las instrucciones de uso.
    MsgText = "Para insertar un nuevo objeto, elija Nuevo en el menú Archivo y seleccione un objeto en el cuadro de diálogo Insertar objeto."
    MsgText = MsgText & PB & "Cuando haya guardado e insertado un objeto usando el comando Guardar como, podrá usar el comando Abrir en el menú Archivo para ver el objeto en las siguientes sesiones."
    MsgText = MsgText & PB & "Para editar un objeto, haga doble clic en el objeto para mostrar el entorno de diseño de la aplicación desde la que se generó el objeto."
    MsgText = MsgText & PB & "Haga clic en el objeto con el botón secundario del mouse para ver los verbos del objeto."
    MsgText = MsgText & PB & "Use los comandos Copiar, Eliminar y Pegado especial para copiar, eliminar y pegar objetos."
    MsgText = MsgText & PB & "Elija el comando Actualizar para actualizar el contenido del objeto insertable."
    MsgBox MsgText, 64, "Instrucciones de la demostración del control contenedor OLE"
End Sub

Sub NewOleForm()
    Dim Newform As New ChildForm
    Newform.Show
    ' Sólo presenta el cuadro de diálogo Insertar objeto si el usuario eligió Nuevo en el menú Archivo.
    If MDINew Then
        ParentForm.ActiveForm.OLE1.InsertObjDlg
    End If
    
    UpdateCaption
End Sub

Sub OpenObject()
    MDINew = False
    NewOleForm
    OpenSave ("Open")
    If ParentForm.ActiveForm.OLE1.OLEType = vbOLENone Then
        Unload ParentForm.ActiveForm
    End If
End Sub

' La apertura de un archivo sólo funciona si el archivo contiene un objeto 
' de Automatización OLE válido.
' Para ver este funcionamiento, siga este procedimiento al ejecutar la aplicación.
' 1) En el menú Archivo, elija Nuevo y especifique un objeto.
' 2) Edite el objeto y después elija Guardar como en el menú Archivo.
' 3) Haga clic en el cuadro de control del objeto para cerrarlo.
' 4) En el menú Archivo, elija Abrir y seleccione el archivo recién guardado.
Sub OpenSave(Action As String)
    Dim Filenum
    Filenum = FreeFile
    ' Establece las opciones y filtros del control CommonDialog.
    ParentForm.ActiveForm.CommonDialog1.Filter = _
      "Objetos insertables (*.OLE)|*.OLE|Todos los archivos (*.*)|*.*"
    ParentForm.ActiveForm.CommonDialog1.FilterIndex = 1
  
    ParentForm.ActiveForm.OLE1.FileNumber = Filenum

On Error Resume Next

    Select Case Action
        Case "Save"
            ' Presenta el cuadro de diálogo Guardar como.
            ParentForm.ActiveForm.CommonDialog1.ShowSave
            If Err Then
                ' El usuario eligió Cancelar.
                If Err = 32755 Then
                    Exit Sub
                Else
                    MsgBox "Error inesperado en el cuadro de diálogo Guardar como."
                End If
            End If
            ' Abre y guarda el archivo.
            Open ParentForm.ActiveForm.CommonDialog1.filename For Binary As Filenum
            If Err Then
                MsgBox (Error)
                    Exit Sub
            End If
                ParentForm.ActiveForm.OLE1.SaveToFile Filenum
            If Err Then MsgBox (Error)

        Case "Open"
            ' Presenta el cuadro de diálogo Abrir archivo.
            ParentForm.ActiveForm.CommonDialog1.ShowOpen
            If Err Then
                ' El usuario eligió Cancelar.
                If Err = 32755 Then
                    Exit Sub
                Else
                    MsgBox "Error inesperado en el cuadro de diálogo Abrir archivo."
                End If
            End If
            ' Abre el archivo.
            Open ParentForm.ActiveForm.CommonDialog1.filename For Binary As Filenum
            If Err Then
                Exit Sub
            End If
            ' Presenta el puntero de reloj de arena.
            Screen.MousePointer = 11
            ParentForm.ActiveForm.OLE1.ReadFromFile Filenum
            If (Err) Then
                If Err = 30015 Then
                    MsgBox "No es un objeto válido."
                Else
                    MsgBox Error$
                End If
                Unload ParentForm.ActiveForm
            End If
            ' Si no hay errores al abrir, activa el objeto.
            ParentForm.ActiveForm.OLE1.DoVerb -1

        ' Establece las propiedades del formulario ahora que el control contenedor OLE contiene un objeto.
        UpdateCaption
        ' Restablece el puntero del mouse.
        Screen.MousePointer = 0
    End Select
  
    Close Filenum
End Sub

Sub UpdateCaption()
    ' Establece las propiedades del formulario ahora que contiene un objeto.
    ParentForm.ActiveForm.Caption = "Objeto " + ParentForm.ActiveForm.OLE1.Class
    On Error Resume Next
End Sub

