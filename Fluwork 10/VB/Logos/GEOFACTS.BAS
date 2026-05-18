Attribute VB_Name = "Module1"
Option Explicit

Public wbWorld As Excel.Workbook
Public shtWorld As Excel.Worksheet

Sub Setup()
    ChDir App.Path
    ChDrive App.Path
    ' Obtiene la primera hoja de WORLD.XLS.
    Set shtWorld = GetObject("world.xls")
    ' Obtiene el libro de trabajo.
    Set wbWorld = shtWorld.Application.Workbooks("world.xls")
End Sub

' Establece los objetos a Nothing.
Sub CleanUp()
    ' Esto fuerza la descarga de Microsoft Excel,
    ' suponiendo que no haya sido cargado por otras aplicaciones u otros usuarios.
    Set shtWorld = Nothing
    Set wbWorld = Nothing
End Sub

' Llena el cuadro combinado Continentes con los nombres
' de las hojas del libro de trabajo.
Sub FillContinentsList()
    Dim shtContinent As Excel.Worksheet
    
    ' Recorre la colección de hojas y agrega
    ' el nombre de cada hoja al cuadro combinado.
    For Each shtContinent In wbWorld.Sheets
        frmGeoFacts.cmbContinents.AddItem shtContinent.Name
    Next
    ' Selecciona el primer elemento y lo presenta en el cuadro combinado.
    frmGeoFacts.cmbContinents.Text = frmGeoFacts.cmbContinents.List(0)

    Set shtContinent = Nothing
End Sub

' Llena el cuadro combinado Continentes con los nombres
' de las características que corresponden a un continente dado.
Sub FillFeaturesList()
    Dim shtContinent As Excel.Worksheet
    Dim rngFeatureList As Excel.Range
    Dim intFirstBlankCell As Integer
    Dim loop1 As Integer

    ' Oculta la lista antigua.
    frmGeoFacts.lstTopRanking.Visible = False
    
    ' Obtiene la hoja cuyo nombre es el la del continente seleccionado en el cuadro combinado Continentes.
    Set shtContinent = wbWorld.Sheets(frmGeoFacts.cmbContinents.Text)
    ' Asigna la primera fila de esta hoja a un objeto.
    Set rngFeatureList = shtContinent.Rows(1)
    
    ' Verifica si es una lista vacía.
    If (rngFeatureList.Cells(1, 1) = "") Then
        intFirstBlankCell = 0
    Else
        ' Busca la primera celda en blanco de la fila.
        intFirstBlankCell = rngFeatureList.Find("").Column
    End If
    
    ' Borra el contenido anterior del cuadro combinado de características del continente.
    frmGeoFacts.cmbFeatures.Clear
            
    ' Agrega elementos al cuadro combinado de características.
    For loop1 = 1 To intFirstBlankCell
            frmGeoFacts.cmbFeatures.AddItem rngFeatureList.Cells(1, loop1)
    Next
    
    ' Selecciona el primer elemento y lo presenta en el cuadro combinado.
    frmGeoFacts.cmbFeatures.Text = frmGeoFacts.cmbFeatures.List(0)

    ' Limpieza.
    Set shtContinent = Nothing
    Set rngFeatureList = Nothing
End Sub

' Llena la lista de elementos por rango.
Sub FillTopRankingList()
    Dim shtContinent As Excel.Worksheet
    Dim intColumOfFeature As Integer
    Dim rngRankedList As Excel.Range
    Dim intFirstBlankCell As Integer
    Dim loop1 As Integer
    
    ' Obtiene la hoja cuyo nombre es el la del continente seleccionado en el cuadro combinado Continentes.
    Set shtContinent = wbWorld.Sheets(frmGeoFacts.cmbContinents.Text)
    
    ' Borra el contenido anterior del cuadro de lista lstTopRanking.
    frmGeoFacts.lstTopRanking.Clear
    
    ' Si la selección de Característica del continente está vacía, no hace nada.
    If (frmGeoFacts.cmbFeatures <> "") Then
        
        ' Busca la columna de la característica seleccionada en la primera fila de la hoja de cálculo.
        intColumOfFeature = shtContinent.Rows(1).Find(frmGeoFacts.cmbFeatures.Text).Column
        
        ' Asigna la columna a un objeto.
         Set rngRankedList = shtContinent.Columns(intColumOfFeature)
        
        ' Verifica si es una lista vacía.
        If (rngRankedList.Cells(1, 1) = "") Then
            intFirstBlankCell = 0
        Else
            ' Busca la primera celda en blanco de la fila.
            intFirstBlankCell = rngRankedList.Find("").Row
        End If
                
        ' Agrega los elementos al cuadro de lista lstTopRanking.
        For loop1 = 2 To intFirstBlankCell
            frmGeoFacts.lstTopRanking.AddItem rngRankedList.Cells(loop1, 1)
        Next
    
        ' Muestra la nueva lista.
        frmGeoFacts.lstTopRanking.Visible = True
    
    End If
    
    ' Limpieza.
    Set shtContinent = Nothing
    Set rngRankedList = Nothing
End Sub
