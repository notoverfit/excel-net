Attribute VB_Name = "Training"
Sub TrainNetwork()
    StopTraining = False

    ' training info
    Dim epochs As Long
    Dim alpha As Double
    Dim currMat As Variant
    
    ' sheets
    Dim networks As Worksheet
    Dim weights As Worksheet
    Dim trainData As Variant: trainData = LoadFeatures("Train_Data")
    Dim targets As Variant: targets = LoadTargets("Train_Data")
    Dim layerInput As Variant
    
    ' network info
    Dim inputSize As Long: inputs = GetInputSize()
    Dim outputSize As Long: outputSize = 1
    Dim layers As Long: layers = NumberOfLayers()
    Dim W() As Variant
    Dim Activations As Collection
    Dim currLoss As Double
    
    ' epoch/loss chart info
    Dim epochCol As Long: epochCol = 7
    Dim lossCol As Long: lossCol = 8
    Dim lossChart As ChartObject: Set lossChart = InitPlot()
    
    
    ' state variables for backprop
    Dim predictions As Variant
    Dim M As Long: M = UBound(trainData, 1)
    
    ' load in worksheets
    Set networks = ThisWorkbook.Worksheets("Network")
    Set weights = ThisWorkbook.Worksheets("Weights")
    
    ' clear epoch/loss history columns
    networks.Columns("G:H").ClearContents
    
    W = LoadWeights()
    Set Activations = LoadActivations()

    epochs = networks.Cells(2, 2).Value
    alpha = networks.Cells(3, 2).Value
    
    Dim preOutput As Collection
    Dim postOutput As Collection

    For epoch = 1 To epochs
        If StopTraining Then Exit Sub
        Set preOutput = New Collection
        Set postOutput = New Collection
        ' forward pass
        For layer = LBound(W) To UBound(W)
            If layer = 1 Then
                layerInput = AddBiasColumn(trainData)
            Else
                layerInput = AddBiasColumn(currMat)
            End If
    
            currMat = MatMult(layerInput, T(W(layer)))
            
            preOutput.Add currMat
            
            ' now add transformations
            If layer = UBound(W) Then
                currMat = SigmoidMatrix(currMat)
            ElseIf Activations(layer) = "ReLu" Then
                currMat = ReLuMatrix(currMat)
            ElseIf Activations(layer) = "Sigmoid" Then
                currMat = SigmoidMatrix(currMat)
            End If
    
            postOutput.Add currMat
        Next layer
        
        ' first calculate dL/dy, the first step of the dz gradient
        
        predictions = postOutput(postOutput.Count)
    
        ' state variables for back-propagation
        Dim dZ As Variant: dZ = MatSubtract(predictions, targets)
        Dim prevA As Variant

        currLoss = loss(predictions, targets)
        networks.Cells(epoch + 1, epochCol).Value = epoch
        networks.Cells(epoch + 1, lossCol).Value = currLoss
        Call UpdatePlot(lossChart, networks, epoch)
        
        For layer = UBound(W) To LBound(W) Step -1
            If layer = 1 Then
                prevA = AddBiasColumn(trainData)
            Else
                prevA = AddBiasColumn(postOutput(layer - 1))
            End If
            
            dW = MatMult(T(dZ), prevA)
            
            ' average the gradient by training samples
            dW = ElementMultiply(dW, 1 / M)
            
            ' now we update dZ
            If layer > LBound(W) Then
                W_NoBias = RemoveBiasColumn(W(layer))
                
                ' (dL / dy) * (dy / dZn) * (dZn / dAn_1) * (dAn_1 / dZn_1)
                If Activations(layer - 1) = "ReLu" Then
                    dZ = MatrixElementMultiply(MatMult(dZ, W_NoBias), dxReLuMatrix(preOutput(layer - 1)))
                ElseIf Activations(layer - 1) = "Sigmoid" Then
                    dZ = MatrixElementMultiply(MatMult(dZ, W_NoBias), dxSigmoidMatrix(preOutput(layer - 1)))
                End If
            End If
            
            ' update weight by the learning step
            W(layer) = MatSubtract(W(layer), ElementMultiply(dW, alpha))
        Next layer
        
        ' save weights so you can see them change live
        SaveWeights (W)
        
        ' updates plot
        DoEvents
    Next epoch
    
    ' report final loss
    MsgBox "[SUCCESS] Training completed. Loss = " & currLoss
End Sub

Public Function NumberOfLayers() As Long
    Dim network As Worksheet
    Dim hiddenLayers As Long
    
    Set network = ThisWorkbook.Worksheets("Network")
    hiddenLayers = network.Cells(network.rows.Count, 5).End(xlUp).Row - 1
    
    NumberOfLayers = hiddenLayers + 2
End Function
