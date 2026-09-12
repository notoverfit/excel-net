Attribute VB_Name = "Loaders"
Public Function LoadFeatures(ByVal sheetName As String) As Variant
    Dim ws As Worksheet
    Dim inputSize As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    inputSize = GetInputSize()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    LoadFeatures = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, inputSize)).Value2
End Function

Public Function LoadTargets(ByVal sheetName As String) As Variant
    Dim ws As Worksheet
    Dim inputSize As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(sheetName)

    ' find range of targets
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    LoadTargets = ws.Range( _
        ws.Cells(2, lastCol), _
        ws.Cells(lastRow, lastCol) _
    ).Value2
End Function


Public Function LoadActivations() As Collection
    Dim network As Worksheet
    Dim Activations As New Collection
    Dim activColumn As Long: activColumn = 6

    Set network = ThisWorkbook.Worksheets("Network")
    lastRow = network.Cells(network.rows.Count, activColumn).End(xlUp).Row
    
    For r = 2 To lastRow
        Activations.Add network.Cells(r, activColumn).Value
    Next r
    
    Set LoadActivations = Activations
End Function

Public Function LoadWeights() As Variant

    Dim network As Worksheet
    Dim wsWeights As Worksheet

    Dim neuronColumn As Long: neuronColumn = 5
    Dim previousNeurons As Long
    Dim layerNeurons As Long

    Dim outputRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long

    Dim currentW As Variant
    Dim weightMatrices() As Variant

    Set network = ThisWorkbook.Worksheets("Network")
    Set wsWeights = ThisWorkbook.Worksheets("Weights")

    previousNeurons = GetInputSize()
    outputRow = 0

    lastRow = network.Cells(network.rows.Count, neuronColumn).End(xlUp).Row

    ' Hidden layers + output layer
    ReDim weightMatrices(1 To lastRow)

    i = 1

    ' Hidden layers
    For r = 2 To lastRow

        layerNeurons = network.Cells(r, neuronColumn).Value2

        currentW = wsWeights.Range( _
            wsWeights.Cells(outputRow + 1, 1), _
            wsWeights.Cells(outputRow + layerNeurons, previousNeurons + 1) _
        ).Value2

        weightMatrices(i) = currentW

        i = i + 1
        outputRow = outputRow + layerNeurons + 1
        previousNeurons = layerNeurons

    Next r

    ' Output layer
    layerNeurons = 1

    currentW = wsWeights.Range( _
        wsWeights.Cells(outputRow + 1, 1), _
        wsWeights.Cells(outputRow + layerNeurons, previousNeurons + 1) _
    ).Value2

    weightMatrices(i) = currentW

    LoadWeights = weightMatrices

End Function

Public Sub SaveWeights(ByVal W As Variant)

    Dim wsWeights As Worksheet
    Dim layer As Long
    Dim outputRow As Long
    Dim rows As Long
    Dim cols As Long
    Dim currentW As Variant

    Set wsWeights = ThisWorkbook.Worksheets("Weights")

    wsWeights.Cells.ClearContents
    outputRow = 0

    For layer = LBound(W) To UBound(W)

        currentW = W(layer)

        rows = UBound(currentW, 1)
        cols = UBound(currentW, 2)

        wsWeights.Range( _
            wsWeights.Cells(outputRow + 1, 1), _
            wsWeights.Cells(outputRow + rows, cols) _
        ).Value2 = currentW

        outputRow = outputRow + rows + 1

    Next layer

End Sub
