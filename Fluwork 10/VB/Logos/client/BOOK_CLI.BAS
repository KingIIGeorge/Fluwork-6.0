Attribute VB_Name = "ClientMain"
Option Explicit

Global goBookSaleSvrInst As Object    ' Referencia al objeto BookSale Server Business
Global goStatusPanel As Object        ' Referencia al Panel de la barra de estado

Global gstrPubRevenue As String
Global gstrAuthorRoyalty As String
Global gstrDBName As String           ' Contiene una cadena que se refiere a la base
                                      ' de datos global

Global gDB As Database                ' Referencia de objeto a la base de datos global
Global gSn As Recordset

Sub Main()

On Error GoTo MainError

  'Aquí se asume que la base de datos de ventas de libros se encuentra en el directorio
  'actual. Consulte el código de form_load para frmRevenue para ver cómo se trata este punto
  gstrDBName = "booksale.mdb"

  Load frmRevenue
  
  Exit Sub
  
MainError:
  MsgBox Error$ & " - " & Str$(Err), vbCritical, "Error al iniciar el cliente de BookSale"
  Exit Sub

End Sub

