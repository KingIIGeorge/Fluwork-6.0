Attribute VB_Name = "def"
Global fichaactual As Long
Global Const pconfig = "c:\estados.dat"
Global BASE As Long
Global dbpath As String
Global touchedreally As Boolean
Global tmpficha As Long
Global Const MAX_CANT_RESULTS = 100
Global unidad As String
Global directorio As String
Global nombrecompleto As String
Global conpre As String * 1
Global showres As Boolean
Global onlyone As Boolean
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global estados(15) As Testados
Global registro As Tregistro
Global regindex As Tindexregistro
Global persona As tpersona
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function enter_a_br(que As String) As String
enter_a_br = TextFormat.EnterToBr(que)
End Function

Function pipe_a_br(que As String) As String
pipe_a_br = TextFormat.PipeToBr(que)
End Function

Function AddBackslash(ByVal str As String) As String
AddBackslash = PathUtils.EnsureTrailingBackslash(str)
End Function

Function getlastfichanumber()
getlastfichanumber = RecordManager.LastFichaNumber()
End Function

Function MostrarFicha(ficha As Long)

Call RecordFormView.PrepararContenedorFicha()

If Not touchedreally Then Exit Function
Call RecordFormView.LimpiarFormularioFicha()
Call RecordManager.ReadFicha(ficha, registro)
Call RecordFormView.PrepararPantallaFichaExistente()
Call RecordFormView.CargarRegistroEnFormulario()

Exit Function

ControlError:   ' Rutina de control de errores.
    Select Case Err.Number  ' Evalua el numero de error.
        Case 55:
                 MsgBox "El archivo ya esta abierto"
                 Close #1
        Exit Function
        Case Else
        MsgBox "Ficha no existente"
        Close #1
        Exit Function
    End Select
    
End Function
