Attribute VB_Name = "RecordFormView"
Option Explicit

Public Sub PrepararContenedorFicha()
Form1.utilizardatos.Enabled = True
Form1.frame2.Visible = False
Form1.Frame1.Width = 11660
Form1.Frame1.Height = 6990
End Sub

Public Sub LimpiarFormularioFicha()
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
End Sub

Public Sub PrepararPantallaFichaExistente()
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
End Sub

Public Sub CargarRegistroEnFormulario()
Dim statusColor As Long

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
Call ActualizarConfirmacionFormulario()
Form1.tdireccion.Text = Trim(registro.direccion)
Form1.temail.Text = Trim(registro.email)
Form1.Tllamareldia.Text = Trim(registro.llamareldia)
Form1.Tcontroladopor.Text = Trim(registro.controladopor)
Form1.Tavisadoeldia.Text = Trim(registro.avisadoeldia)
Form1.Tavisadopor.Text = Trim(registro.avisadopor)
statusColor = StatusCatalog.ColorForStatus(Form1.lbllista.Caption)
If statusColor <> -1 Then Form1.lbllista.ForeColor = statusColor
End Sub

Private Sub ActualizarConfirmacionFormulario()
If Mid$(Form1.tconfirmacion.Text, 1, 2) = "N-" Then
Form1.Label16.Caption = "NO CONFIRMADO"
ElseIf Mid$(Form1.tconfirmacion.Text, 1, 2) = "C-" Then
Form1.Label16.Caption = "CONFIRMADO"
Else
Form1.Label16.Caption = "NO DISPONIBLE"
End If
End Sub
