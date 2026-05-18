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
Dim resultado As String
Dim i As Long
Dim pos As Long
resultado = ""
pos = 0

For i = 1 To Len(que)
If Mid(que, i, 1) = vbCr Or i = Len(que) Then
If pos = 0 Then
resultado = Mid(que, 1, i - 1) & "<br>"
ElseIf i = Len(que) Then
resultado = resultado & Mid(que, pos + 1, i - pos) & "<br>"
Else
resultado = resultado & Mid(que, pos + 1, i - pos - 1) & "<br>"
End If
pos = i + 1
End If
Next i
enter_a_br = resultado
End Function


Function pipe_a_br(que As String) As String
Dim resultado As String
Dim i As Long
Dim pos As Long
Dim npipe As Byte

npipe = 0
resultado = ""
pos = 0

For i = 1 To Len(que)
If Mid(que, i, 1) = "|" Or i = Len(que) Then
npipe = npipe + 1
    If pos = 0 Then
    resultado = Mid(que, 1, i - 1) & "   "
    ElseIf i = Len(que) Then
    resultado = resultado & Mid(que, pos, i - pos + 1) & "<BR>"
    Else
    If npipe Mod 2 = 0 Then
    resultado = resultado & Mid(que, pos, i - pos) & "<BR>"
    Else
    resultado = resultado & Mid(que, pos, i - pos) & "   "
    End If
    End If
    pos = i + 1
End If
Next i
pipe_a_br = resultado
End Function

Function AddBackslash(ByVal str As String) As String
    If (Right(str, 1) = "\") Then
        AddBackslash = str
    Else
        AddBackslash = str & "\"
End If
End Function

Function getlastfichanumber()
Dim archivo As String
Dim tam As Long
Dim reg2 As Tindexregistro
archivo = Trim(dbpath + "\index.dat")
Open archivo For Random Access Read As #138 Len = Len(reg2)
tam = LOF(138)
If tam > 0 Then
Get #138, tam / Len(reg2), reg2
getlastfichanumber = Val(Trim(reg2.ficha))
Else
getlastfichanumber = BASE
End If
Close #138
End Function

Function MostrarFicha(ficha As Long)

Form1.utilizardatos.Enabled = True
Form1.frame2.Visible = False
Form1.Frame1.Width = 11660
Form1.Frame1.Height = 6990

If Not touchedreally Then Exit Function
Form1.tficha.Text = ""
Form1.tfullname.Text = ""
Form1.tprecio.Text = ""
Form1.tpresupuesto.Text = ""
Form1.tproblema.Text = ""
Form1.tsolucion.Text = ""
Form1.tadjuntos.Text = ""
Form1.ttelefono.Text = ""
Form1.tmodelo.Text = ""
Form1.tnserie.Text = ""
Form1.tfechaingreso.Text = ""
Form1.tfechaegreso.Text = ""
Form1.ttecnico.Text = ""
Form1.tatendidopor.Text = ""
Form1.lbllista.Caption = ""
Form1.tconfirmacion.Text = ""
Form1.tdireccion.Text = ""
Form1.temail.Text = ""
Form1.Tllamareldia.Text = ""
Form1.Tcontroladopor.Text = ""
Form1.Tavisadoeldia.Text = ""
Form1.Tavisadopor.Text = ""

Dim archivo As String
archivo = Trim(dbpath + "\" + Trim(str(ficha)))
Open archivo For Binary Access Read As #4 Len = Len(registro)
Get #4, , registro
Close #4

If showres = True Then
Form1.Command12.BackColor = QBColor(14)
Else
Form1.Command12.BackColor = QBColor(8)
End If

If onlyone = True Then
Form1.Command12.Visible = False
Form1.Command12.Enabled = False
Else
Form1.Command12.Visible = True
Form1.Command12.Enabled = True
End If

Form1.cmdnuevo.Visible = False
Form1.cmdnuevo.Enabled = False
Form1.Frame1.Visible = True
Form1.tfechaingreso.Enabled = True
Form1.tfechaegreso.Enabled = True
Form1.ttecnico.Enabled = True
Form1.tatendidopor.Enabled = True
Form1.tmodelo.Enabled = True
Form1.tmodelo.Visible = True
Form1.tnserie.Enabled = True
Form1.tnserie.Visible = True

Form1.cmdgrabar.Visible = True
Form1.cmdgrabar.Enabled = True
Form1.cmdcancel.Enabled = True
Form1.cmdcancel.Visible = True
Form1.cmdimprimir.Visible = True
Form1.cmdimprimir.Enabled = True
Form1.cmdprintpublic.Enabled = True
Form1.cmdprintpublic.Visible = True
Form1.Command11.Enabled = True
Form1.Command11.Visible = True
Form1.tficha.Enabled = True
Form1.tfullname.Enabled = True
Form1.tprecio.Enabled = True
Form1.tpresupuesto.Enabled = True
Form1.tproblema.Enabled = True
Form1.tadjuntos.Enabled = True
Form1.tsolucion.Enabled = True
Form1.ttelefono.Enabled = True
Form1.tconfirmacion.Enabled = True
Form1.tdireccion.Enabled = True
Form1.temail.Enabled = True
Form1.Tllamareldia.Enabled = True
Form1.Tcontroladopor.Enabled = True
Form1.Tavisadoeldia.Enabled = True
Form1.Tavisadopor.Enabled = True

Form1.lbllista.Caption = registro.estado

Form1.tficha.SetFocus

Form1.tficha.Text = Trim(registro.ficha)
Form1.tfullname.Text = Trim(registro.fullname)
Form1.tprecio.Text = Trim(registro.precio)
Form1.tpresupuesto.Text = Trim(registro.presupuesto)
Form1.tproblema.Text = Trim(registro.problema)
Form1.tsolucion.Text = Trim(registro.solucion)
Form1.tadjuntos.Text = Trim(registro.adjuntos)
Form1.ttelefono.Text = Trim(registro.telefono)
Form1.tfechaingreso.Text = Trim(registro.fechaingreso)
Form1.tfechaegreso.Text = Trim(registro.fechaegreso)
Form1.ttecnico.Text = Trim(registro.tecnico)
Form1.tatendidopor.Text = Trim(registro.atendidopor)
Form1.tnserie.Text = Trim(registro.nserie)
Form1.tmodelo.Text = Trim(registro.modelo)
Form1.tconfirmacion.Text = Trim(registro.confirmacion)
If Mid$(Form1.tconfirmacion.Text, 1, 2) = "N-" Then
Form1.Label16.Caption = "NO CONFIRMADO"
ElseIf Mid$(Form1.tconfirmacion.Text, 1, 2) = "C-" Then
Form1.Label16.Caption = "CONFIRMADO"
Else
Form1.Label16.Caption = "NO DISPONIBLE"
End If
Form1.tdireccion.Text = Trim(registro.direccion)
Form1.temail.Text = Trim(registro.email)
Form1.Tllamareldia.Text = Trim(registro.llamareldia)
Form1.Tcontroladopor.Text = Trim(registro.controladopor)
Form1.Tavisadoeldia.Text = Trim(registro.avisadoeldia)
Form1.Tavisadopor.Text = Trim(registro.avisadopor)
If Trim(Form1.lbllista.Caption) = Trim("POR VER") Then
Form1.lbllista.ForeColor = QBColor(11)
End If
If Trim(Form1.lbllista.Caption) = Trim("REPARANDO") Then
Form1.lbllista.ForeColor = QBColor(12)
End If
If Trim(Form1.lbllista.Caption) = Trim("LISTA") Then
Form1.lbllista.ForeColor = QBColor(10)
End If
If Trim(Form1.lbllista.Caption) = Trim("ENTREGADA") Then
Form1.lbllista.ForeColor = QBColor(8)
End If
If Trim(Form1.lbllista.Caption) = Trim("STD/BY") Then
Form1.lbllista.ForeColor = QBColor(13)
End If
If Trim(Form1.lbllista.Caption) = Trim("CHEQUEO") Then
Form1.lbllista.ForeColor = QBColor(14)
End If
If Trim(Form1.lbllista.Caption) = Trim("REP.EXT.") Then
Form1.lbllista.ForeColor = QBColor(12)
End If
If Trim(Form1.lbllista.Caption) = Trim("PV EXT.") Then
Form1.lbllista.ForeColor = QBColor(11)
End If
If Trim(Form1.lbllista.Caption) = Trim("LISTA NR") Then
Form1.lbllista.ForeColor = QBColor(10)
End If
If Trim(Form1.lbllista.Caption) = Trim("LISTA BRGS") Then
Form1.lbllista.ForeColor = QBColor(10)
End If
If Trim(Form1.lbllista.Caption) = Trim("PRESUP") Then
Form1.lbllista.ForeColor = QBColor(15)
End If
If Trim(Form1.lbllista.Caption) = Trim("ENTREGAR") Then
Form1.lbllista.ForeColor = QBColor(9)
End If
If Trim(Form1.lbllista.Caption) = Trim("ANULADA") Then
Form1.lbllista.ForeColor = QBColor(2)
End If
If Trim(Form1.lbllista.Caption) = Trim("DEPOSITO") Then
Form1.lbllista.ForeColor = QBColor(8)
End If
If Trim(Form1.lbllista.Caption) = Trim("DIAGNOSTIC") Then
Form1.lbllista.ForeColor = QBColor(14)
End If

Exit Function

ControlError:   ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
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

