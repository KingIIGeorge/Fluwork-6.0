VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Ejemplo de DBGrid no enlazado"
   ClientHeight    =   3345
   ClientLeft      =   2325
   ClientTop       =   1815
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5910
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3240
      Left            =   60
      OleObjectBlob   =   "UnBndGrd.frx":0000
      TabIndex        =   0
      Top             =   60
      Width           =   5820
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Base = "0{8BACA63C-DF27-11CF-B1A5-00AA003B0266}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Ejemplo de DBGrid no enlazado
'
' El propósito de este ejemplo es darle una idea de lo que tiene que hacer
' para que el control DBGrid funcione en modo no enlazado (DataMode = 1).
' Por simplificar, este ejemplo utiliza una matriz llamada UserData para
' almacenar los datos para la cuadrícula. En la mayoría de los casos, el
' control DBGrid se puede enlazar a un control Data; sin embargo, algunas
' situaciones pueden requerir que lo utilice en modo no enlazado, como
' con un formato de base de datos patentado o para simples conjuntos de
' datos para los que el control Data no ofrece compatibilidad.

Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 3 ' Número máximo de campos del conjunto de registros.

' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr 
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf 
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 2, 0 To 14)

' Fila 1
UserData(0, 0) = "Cristina"
UserData(1, 0) = "Martínez"
UserData(2, 0) = "19"

' Fila 2
UserData(0, 1) = "Juan"
UserData(1, 1) = "Pardo"
UserData(2, 1) = "12"
' Fila 3
UserData(0, 2) = "Mara"
UserData(1, 2) = "Belén"
UserData(2, 2) = "33"
' Fila 4
UserData(0, 3) = "Daniel"
UserData(1, 3) = "Rendich"
UserData(2, 3) = "40"
' Fila 5
UserData(0, 4) = "Federico"
UserData(1, 4) = "Couto"
UserData(2, 4) = "42"
' Fila 6
UserData(0, 5) = "María Asunción"
UserData(1, 5) = "Rodríguez"
UserData(2, 5) = "28"
' Fila 7
UserData(0, 6) = "Agustina"
UserData(1, 6) = "Rivera"
UserData(2, 6) = "29"
' Fila 8
UserData(0, 7) = "Enrique"
UserData(1, 7) = "Ballinea"
UserData(2, 7) = "42"
' Fila 9
UserData(0, 8) = "Susana"
UserData(1, 8) = "Medrano"
UserData(2, 8) = "31"
' Fila 10
UserData(0, 9) = "Guillermo"
UserData(1, 9) = "Alzaga"
UserData(2, 9) = "30"
' Fila 11
UserData(0, 10) = "Luis"
UserData(1, 10) = "Romero"
UserData(2, 10) = "28"
' Fila 12
UserData(0, 11) = "Ernesto"
UserData(1, 11) = "Méndez"
UserData(2, 11) = "60"
' Fila 13
UserData(0, 12) = "Antonio"
UserData(1, 12) = "Abelló"
UserData(2, 12) = "33"
' Fila 14
UserData(0, 13) = "Ana"
UserData(1, 13) = "Martínez"
UserData(2, 13) = "30"
' Fila 15
UserData(0, 14) = "Francisco"
UserData(1, 14) = "Colusso"
UserData(2, 14) = "60"

mTotalRows& = 15


Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
    DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 2
    DBGrid1.Columns.Add newcnt
    DBGrid1.Columns(newcnt).Caption = "Columna " & newcnt
    DBGrid1.Columns(newcnt).Visible = True
    newcnt = newcnt + 1
Next i

End Sub
