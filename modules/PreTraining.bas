Attribute VB_Name = "PreTraining"
Sub GenerateWeights()
    Randomize

    ' we pull information about hidden layers from the network setup column
    Dim network As Worksheet
    Dim weights As Worksheet
    Dim neuronColumn As Long: neuronColumn = 5
    
    Dim inputNeurons As Long
    Dim previousNeurons As Long
    Dim scaler As Double
    
    ' used for keeping track of which row the weights are at
    Dim outputRow As Long
    
    inputNeurons = GetInputSize()
    previousNeurons = inputNeurons
    
    ' network info is stored in column E; atleast, what is relevant
    Set network = ThisWorkbook.Worksheets("Network")
    Set weights = ThisWorkbook.Worksheets("Weights")
    
    outputRow = 0
    
    lastRow = network.Cells(network.rows.Count, neuronColumn).End(xlUp).Row
    
    weights.Cells.ClearContents
    
    ' now iterate through layers and create matrices
    For r = 2 To lastRow
        layerNeurons = network.Cells(r, neuronColumn).Value

        scaler = Sqr(2# / previousNeurons)
        Call WriteMatrix(weights, outputRow, previousNeurons + 1, layerNeurons, scaler)
        
        ' shift down output row
        outputRow = outputRow + layerNeurons + 1
        
        previousNeurons = layerNeurons
    Next r
    
    ' output layer
    layerNeurons = 1
    
    ' xavier scaling
    scaler = Sqr(layerNeurons / previousNeurons)
    Call WriteMatrix(weights, outputRow, previousNeurons + 1, layerNeurons, scaler)
    
    MsgBox "[SUCCESS] Wrote weight matrices to 'Weights'"
End Sub

Public Function GetInputSize() As Long
    Dim ws As Worksheet
    Dim lastCol As Long
    
    Set ws = ThisWorkbook.Worksheets("Train_Data")
    
    ' this function goes to the last column, first row possible (row, column)
    ' then, it does ctrl + left arrow (.End(xlToLeft)), and gets the column number
    ' excluding the class, this must be the neuron length
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    GetInputSize = lastCol - 1
End Function

Public Sub WriteMatrix(ByVal ws As Worksheet, ByVal startRow As Long, ByVal inputCount As Long, ByVal outputCount As Long, ByVal scaler As Double)
    Dim i As Long, j As Long

    For i = 1 To outputCount
        For j = 1 To inputCount
            If j = inputCount Then
                ws.Cells(startRow + i, j).Value = 0#
            Else
                ws.Cells(startRow + i, j).Value = RandomNormal() * scaler
            End If
        Next j
    Next i
End Sub

Public Function RandomNormal() As Double
    Dim u1 As Double, u2 As Double
    u1 = Rnd()
    If u1 = 0 Then u1 = 0.0000001
    u2 = Rnd()
    RandomNormal = Sqr(-2# * Log(u1)) * Cos(2# * WorksheetFunction.Pi() * u2)
End Function
