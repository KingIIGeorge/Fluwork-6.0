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
Type tpersona
    persona1 As String * 30
    persona2 As String * 30
    persona3 As String * 30
    persona4 As String * 30
    persona5 As String * 30
    persona6 As String * 30
    persona7 As String * 30
    persona8 As String * 30
    persona9 As String * 30
    persona10 As String * 30
    persona11 As String * 30
    persona12 As String * 30
    persona13 As String * 30
    
    persona14 As String * 200
    persona15 As String * 200
    persona16 As String * 200
    persona17 As String * 200
    persona18 As String * 200
    persona19 As String * 200
    persona20 As String * 200
    persona21 As String * 200
    persona22 As String * 200
    persona23 As String * 200
    persona24 As String * 200
    persona25 As String * 200
    persona26 As String * 200
    
    
    End Type

Type Tregistro
    ficha As String * 10
    fechaingreso As String * 10
    fechaegreso As String * 10
    estado As String * 10
    fullname As String * 50
    telefono As String * 15
    adjuntos As String * 1024
    problema As String * 1024
    solucion As String * 1024
    presupuesto As String * 10
    precio As String * 10
    atendidopor As String * 50
    tecnico As String * 50
    modelo As String * 50
    nserie As String * 50
    direccion As String * 200
    email As String * 75
    llamareldia As String * 30
    controladopor As String * 50
    avisadoeldia As String * 30
    avisadopor As String * 30
    confirmacion As String * 30
End Type

Type Tindexregistro
    ficha As String * 10
    fullname As String * 50
    telefono As String * 15
    modelo As String * 50
    fecha As String * 10
    estado As String * 10
    tecnico As String * 50
    confirmacion As String * 30
        
End Type

Type Testados
    txt As String
End Type

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
