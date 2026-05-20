VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reultados de Busqueda"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9645
   Icon            =   "resu.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483626
      BackColorSel    =   -2147483632
      BackColorBkg    =   12632256
      GridColor       =   8421504
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   4
      FormatString    =   $"resu.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   9480
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Cantidad de fichas encontradas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   7080
      Width           =   855
   End
   Begin VB.Menu mnuordenar 
      Caption         =   "&Ordenar"
      Begin VB.Menu mnuficha 
         Caption         =   "Ficha"
      End
      Begin VB.Menu mnufecha 
         Caption         =   "Fecha"
      End
      Begin VB.Menu mnunombre 
         Caption         =   "Nombre"
      End
      Begin VB.Menu mnuestado 
         Caption         =   "Estado"
      End
   End
   Begin VB.Menu mnuprint 
      Caption         =   "&Imprimir"
      Begin VB.Menu mnulista 
         Caption         =   "Lista"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
mnuficha.Checked = True
AppState.SetFichaTouched False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
AppState.ClearSearchResults
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
Form1.utilizardatos.Enabled = False
End If
End Sub


Private Sub mnuestado_Click()
If mnufecha.Checked = True Then
mnufecha.Checked = False
End If
If mnuficha.Checked = True Then
mnuficha.Checked = False
End If
If mnunombre.Checked = True Then
mnunombre.Checked = False
End If
Form3.MSFlexGrid1.Col = 5
Form3.MSFlexGrid1.Sort = 1
mnuestado.Checked = True
End Sub

Private Sub mnufecha_Click()
If mnuestado.Checked = True Then
mnuestado.Checked = False
End If
If mnuficha.Checked = True Then
mnuficha.Checked = False
End If
If mnunombre.Checked = True Then
mnunombre.Checked = False
End If
Form3.MSFlexGrid1.Col = 1
Form3.MSFlexGrid1.Sort = 1
mnufecha.Checked = True
End Sub

Private Sub mnuficha_Click()
If mnufecha.Checked = True Then
mnufecha.Checked = False
End If
If mnuestado.Checked = True Then
mnuestado.Checked = False
End If
If mnunombre.Checked = True Then
mnunombre.Checked = False
End If
Form3.MSFlexGrid1.Col = 0
Form3.MSFlexGrid1.Sort = 1
mnuficha.Checked = True
End Sub

Private Sub mnulista_Click()

On Error Resume Next

Dim returnvalue, e
Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro

Form1.Combo1.ListIndex = 0
cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #111 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE

Kill "c:\repext.html"
Close #117

Open "c:\repext.html" For Output As #117

Print #117, ""
Print #117, "<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%""><tr><td width=""25%"" rowspan=""3""><center><IMG border=""0"" SRC=""file:"
Print #117, Trim(dbpath) + "\mag.logo.gif"
Print #117, """ width=""190"" height=""159""></center></td><td width=""25%"" valign=""top"" align=""left"" height=""2""><img border=""0"" src=""file:"
Print #117, Trim(dbpath) + "\direccion.gif"
Print #117, """ width=""200"" height=""69""></td><td width=""25%"" valign=""middle"" align=""center""><p align=""right""><img border=""0"" src=""file:"
Print #117, Trim(dbpath) + "\epson.logo.gif"
Print #117, """ width=""128"" height=""40""></td><td width=""25%"" valign=""middle"" align=""center""><img border=""0"" src=""file:"
Print #117, Trim(dbpath) + "\hp.logo.gif"

Print #117, """ width=""111"" height=""115"">"
Print #117, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" height=""1"" colspan=""3""><p align=""center""><b><font size=""2"">NOTA DE ENVIO DE MERCADERIA A <BR>"

If Form1.Combo2.Text = Form5.Text1.Text Then
Print #117, Trim(Form5.Text13.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text2.Text Then
Print #117, Trim(Form5.Text14.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text3.Text Then
Print #117, Trim(Form5.Text15.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text4.Text Then
Print #117, Trim(Form5.Text16.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text5.Text Then
Print #117, Trim(Form5.Text17.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text6.Text Then
Print #117, Trim(Form5.Text18.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text7.Text Then
Print #117, Trim(Form5.Text19.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text8.Text Then
Print #117, Trim(Form5.Text20.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text9.Text Then
Print #117, Trim(Form5.Text21.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text10.Text Then
Print #117, Trim(Form5.Text22.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text11.Text Then
Print #117, Trim(Form5.Text23.Text)
Else
End If
If Form1.Combo2.Text = Form5.Text12.Text Then
Print #117, Trim(Form5.Text24.Text)
Else
End If

Print #117, "</font></b></p>"
Print #117, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" colspan=""3"">&nbsp;</td></tr></table>"

For i = cantdefichas To 1 Step -1

Get #111, i, rsindex

If InStr(1, Trim(rsindex.estado), Trim(Form1.Command5.Tag)) And InStr(1, Trim(rsindex.tecnico), Trim(Form1.Combo2.Text)) > 0 Then
If cantres < Trim(Form1.tce2.Text) Then
cantres = cantres + 1

Print #117, "<table width= ""100%"" border =""1"" cellspacing =0> <tr> <td width=""5%"">"
Print #117, str(rsindex.ficha)
Print #117, "</td> <td width=""5%"" >"
Print #117, rsindex.fecha

Print #117, "</td> <td width=""30%"">"
Print #117, rsindex.modelo
Print #117, "</td> <td width=""5%"">"
Print #117, rsindex.estado
Print #117, "</td> </tr>"
Else

End If

End If

Next i

Close #111
Print #117, "</table></html>"
Close #117

returnvalue = Shell("c:\Archivos de Programa\Internet Explorer\iexplore.exe C:\repext.html", vbMaximizedFocus)
Unload Form3
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True

End Sub

Private Sub mnunombre_Click()
If mnufecha.Checked = True Then
mnufecha.Checked = False
End If
If mnuficha.Checked = True Then
mnuficha.Checked = False
End If
If mnuestado.Checked = True Then
mnuestado.Checked = False
End If
Form3.MSFlexGrid1.Col = 2
Form3.MSFlexGrid1.Sort = 1
mnunombre.Checked = True
End Sub

Private Sub MSFlexGrid1_Click()

On Error Resume Next
AppState.SelectFicha Val(str(MSFlexGrid1.Text))
Me.Hide
MostrarFicha (tmpficha)
End Sub
