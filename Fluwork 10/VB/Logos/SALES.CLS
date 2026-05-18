VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim sngPubRevenue() As Single     ' Matriz de ingresos mensuales
Dim sngAuthorRoyalty() As Single  ' Matriz de derechos de autor mensuales
Dim sngBookPrice() As Single      ' Matriz de precios de libros mensuales

Public Function GetAuthors() As Variant
'Esta rutina devuelve la lista completa de autores conocidos. Se ofrece como servicio a los
'clientes para que no tengan que saber cómo o dónde obtener la información por sí mismos. Por lo
'tanto, se comporta como un "servicio de datos". Normalmente, estos servicios se agruparían
'separados de los servicios de negocios para evitar dependencias en su desarrollo y mantenimiento.

End Function

Public Function GetBooks(rstrAuthor As String) As Variant
'Esta rutina devuelve la lista de libros que ha publicado el autor especificado. Se ofrece como servicio
'a los clientes para que no tengan que saber cómo o dónde obtener la información por sí mismos. Como tal,
'se comporta como un "servicio de datos". Normalmente, estos servicios se agruparían separados de los
'servicios de negocios para evitar dependencias en su desarrollo y mantenimiento.

End Function

Public Function GetRevenue(intSalesModel As Integer, _
                              curCostPerUnit As Currency, _
                              curAdvCost As Currency, _
                              intSalesPeriod As Integer, _
                              lngUnitsPerMonth As Long, _
                              bolIsDiscount As Boolean, _
                              strBookTitle As String) As Variant

  Dim i As Integer
  Dim iOldBound As Integer
  Dim iNewBound As Integer

  gintSalesModel = intSalesModel
  gcurCostPerUnit = curCostPerUnit
  gcurAdvertisingCost = curAdvCost
  gintSalesPeriod = intSalesPeriod
  glngUnitsPerMonth = lngUnitsPerMonth
  
  If GetPubRevenue(strBookTitle) = False Then
      ServerMsg Error$ & " - " & Str$(Err), vbOKOnly, "Error en GetChartData"
    GetRevenue = 0
    Exit Function
  
  End If

  If GetAuthorRoyalty() = False Then
      ServerMsg Error$ & " - " & Str$(Err), vbOKOnly, "Error en GetChartData"
    GetRevenue = 0
    Exit Function

  End If

  iOldBound = UBound(sngPubRevenue)

  For i = 0 To iOldBound
    sngPubRevenue(i, 1) = sngAuthorRoyalty(i)
  
  Next i

  GetRevenue = sngPubRevenue()

End Function

Public Function GetAuthorRoyalty() As Boolean
Dim i As Integer

Dim cGrossMonthlySalary As Currency
Dim cTaxAmount As Currency
Dim cTotalRevenue As Currency

'Crea una referencia a la clase Tax
Dim objTax As New Taxes

frmBookSales.lblStatus(1).Caption = "Solicitando derechos de autor..."

ReDim sngAuthorRoyalty(gintSalesPeriod)

  For i = 0 To (gintSalesPeriod - 1)
    cGrossMonthlySalary = sngPubRevenue(i, 0) * gRoyalty
    sngAuthorRoyalty(i) = cGrossMonthlySalary - _
                                    objTax.CalcNationalIncomeTax(cGrossMonthlySalary) - _
                                    objTax.CalcSalesTax(cGrossMonthlySalary, 0)

  Next i

  ' Elimina la referencia a la clase
  Set objTax = Nothing
  
  frmBookSales.lblStatus(1).Caption = "Calculando derechos de autor..."
  GetAuthorRoyalty = True

End Function

Public Function GetPubRevenue(strTitle As String) As Variant
Dim sn As Recordset
Dim strSQL As String
Dim i As Integer
Dim Price As Currency

' Crea una referencia a clase
Dim objModel As New Model

Static strOldTitle As String
Static cUnitPrice As Currency

frmBookSales.lblStatus(0).Caption = "Solicitando beneficios del editor..."
frmBookSales.lblStatus(1).Caption = "Calculando beneficios del editor..."

On Error GoTo GetRevenueError

If strTitle <> strOldTitle Then
  frmBookSales.lblStatus(1).Caption = "Buscando fila " & strTitle & "..."
  strSQL = "SELECT Titles.Price " & _
                "FROM Titles " & _
                "WHERE ((Titles.Title=" & Chr$(34) & strTitle & Chr$(34) & "));"
  
  Set sn = gDB.OpenRecordset(strSQL, dbOpenSnapshot)
  cUnitPrice = sn.Fields("Price")
  
Else
  frmBookSales.lblStatus(1).Caption = "Usando los últimos valores..."
  
End If

ReDim sngPubRevenue(gintSalesPeriod - 1, 1)
ReDim sngBookPrice(gintSalesPeriod - 1)

For i = 0 To gintSalesPeriod - 1
   sngPubRevenue(i, 0) = cUnitPrice * _
                          objModel.intGetMonthSales(i, _
                          gintSalesPeriod, _
                          gintSalesModel)
Next i

' Elimina la referencia a clase
Set objModel = Nothing

frmBookSales.lblStatus(1).Caption = "Enviando los benfecios del editor al cliente..."
GetPubRevenue = True

' No intenta cerrar el objeto si no se ha creado la instantánea.
' sn no está nunca definido cuando strTitle = strOldTitle.
  If strTitle <> strOldTitle Then
    sn.Close
    Set sn = Nothing
  End If

  strOldTitle = strTitle

  Exit Function
  If IsObject(sn) Then sn.Close
  Set sn = Nothing

GetRevenueError:
  frmBookSales.lblStatus(1).Caption = Error$ & " - " & Str$(Err)
  GetPubRevenue = False
  
End Function

Private Sub Class_Initialize()
  
On Error GoTo InitErr
  
  If gintInstanceCount = 0 Then
    frmBookSales.Show
    gintInstanceCount = 0
    gDBName = App.Path & "\booksale.mdb"
    
    frmBookSales.lblStatus(1).Caption = "Creando espacio de trabajo"
    DoEvents
    
    Set gWkspc = Workspaces(0)
    
    frmBookSales.lblStatus(1).Caption = "Abriendo " & gDBName & "..."
    Set gDB = gWkspc.OpenDatabase(gDBName, False)
    frmBookSales.lblStatus(1).Caption = "Esperando comando..."
  End If
    
  gintInstanceCount = gintInstanceCount + 1
  frmBookSales.lblInstanceCount.Caption = Format$(gintInstanceCount)

  Exit Sub

InitExit:
Screen.MousePointer = vbDefault
Exit Sub
  
InitErr:
  
  frmBookSales.lblStatus(1).Caption = Error$ & " - " & Str$(Err)
  
  ' Es necesario un tratamiento de errores avanzado cuando no se encuentra
  ' la base de datos, como indicaría el error 3024.
  If Err = 3024 Then
      ' Establece las propiedades de CommonDialog antes de mostrarlo
      With frmBookSales.CommonDialog1
        .DialogTitle = "Imposible encontrar la ubicación del archivo booksale.mdb"
        .Filter = "(*.mdb)|*.mdb"
        .InitDir = CurDir
        .filename = ""
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
        .ShowOpen
        ' Se asegura de que el nombre de archivo no sea una cadena vacía
        If .filename <> "" Then
          ' Se asegura de que el archivo de base de datos obtenido sea realmente
          ' booksale.mdb
          If Right(UCase(.filename), Len("booksale.mdb")) = "BOOKSALE.MDB" Then
            gDBName = .filename
            Set gWkspc = Workspaces(0)
          End If
          Resume
        End If
      End With
  ElseIf Err <> 0 Then ' Otro error
      ServerMsg Error$ & " - " & Str$(Err), vbCritical, "Error al iniciar el servidor BookSale"
      End
  End If
  
  Resume InitExit
  
End Sub



Private Sub Class_Terminate()
  gintInstanceCount = gintInstanceCount - 1
  frmBookSales.lblInstanceCount.Caption = Format$(gintInstanceCount)
  
  If gintInstanceCount = 0 Then
    Unload frmBookSales
  End If
End Sub


