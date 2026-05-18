Attribute VB_Name = "VCRModule"
'*************************************************************
' Propósito: Módulo general de la aplicación de ejemplo VCR.
' Contiene procedimientos compartidos
'*************************************************************
Option Explicit

' Instancia la clase Recorder
Public Recorder As New clsRecorder

'*************************************************************
' Propósito:  Activa o desactiva los botones del
'           formulario VCR según el modo actual.
' Entradas:   Button: el botón de comando que llama
'           al procedimiento.
'*************************************************************
Sub ButtonManager(Button As Control)
    Dim vntControl As Variant   ' Valor de Control
    
    ' determina el botón de función presionado
    ' y actualiza todos los botones y la clase Recorder
    Select Case Button
        Case frmVCR.cmdPlay, frmVCR.cmdForward, frmVCR.cmdRewind
            frmVCR.cmdPause.Enabled = True
            frmVCR.cmdStop.Enabled = True
            frmVCR.cmdRec.Enabled = False
            frmVCR.cmdDown.Enabled = False
            frmVCR.cmdUp.Enabled = False
            Recorder.Enabled = False
        Case frmVCR.cmdRec
            frmVCR.cmdPause.Enabled = True
            frmVCR.cmdStop.Enabled = True
            frmVCR.cmdPlay.Enabled = False
            frmVCR.cmdForward.Enabled = False
            frmVCR.cmdRewind.Enabled = False
            frmVCR.cmdDown.Enabled = False
            frmVCR.cmdUp.Enabled = False
            Recorder.Enabled = False
        Case frmVCR.cmdPause
            frmVCR.cmdPause.Enabled = False
        Case frmVCR.cmdStop
            frmVCR.cmdStop.Enabled = False
            frmVCR.cmdPause.Enabled = False
            frmVCR.cmdRec.Enabled = True
            frmVCR.cmdPlay.Enabled = True
            frmVCR.cmdForward.Enabled = True
            frmVCR.cmdRewind.Enabled = True
            frmVCR.cmdDown.Enabled = True
            frmVCR.cmdUp.Enabled = True
            Recorder.Enabled = True
    End Select
    ' asigna el botón a su botón forma
    ' correspondiente
    vntControl = Button.Name
    vntControl = Right$(vntControl, Len(vntControl) - 3)
    vntControl = "shp" & vntControl
    ' llama a la subrutina que actualiza los indicadores
    HighlightButton vntControl
End Sub
'*************************************************************
' Propósito:  Activa o desactiva los controles forma del
'           formulario VCR según el modo actual.
' Entradas:   MyControl: el control Shape que se pasa al
'           procedimiento.
'*************************************************************
Sub HighlightButton(MyControl As Variant)
    Dim i As Integer    ' Variable contador
    
    ' Recorre la colección Controls
    For i = 0 To frmVCR.Controls.Count - 1
        ' Comprueba si es un control Shape
        If TypeOf frmVCR.Controls(i) Is Shape Then
            ' Si es el control recibido, lo hace
            ' visible, si no, lo hace invisible
            If frmVCR.Controls(i).Name = MyControl Then
                frmVCR.Controls(i).Visible = True
            Else
                frmVCR.Controls(i).Visible = False
            End If
        End If
    Next
End Sub
'*************************************************************
' Propósito:  Almacena el último número de canal para
'           poder restablecerlo después de grabar.
' Entradas:   Channel: el número de canal pasado al
'           procedimiento.
' Devuelve:  El canal almacenado en la variable Static.
'*************************************************************
Function SaveChannel(Channel As Variant) As Integer
    Static intChannel As Integer 'Almacena el canal
    
    ' asigna el valor a la variable Static
    If Channel <> 0 Then
        intChannel = Channel
    End If
    
    ' Devuelve el número del canal
    SaveChannel = intChannel
End Function
