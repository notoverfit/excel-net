Attribute VB_Name = "Testing"
Sub EvaluateNetwork()
    Dim testX As Variant: testX = LoadFeatures("Test_Data")
    Dim testY As Variant: testY = LoadTargets("Test_Data")
    Dim Activations As Collection: Set Activations = LoadActivations()
    Dim W() As Variant: W = LoadWeights()
    
    Dim currMat As Variant
    
    For layer = LBound(W) To UBound(W)
        If layer = 1 Then
            layerInput = AddBiasColumn(testX)
        Else
            layerInput = AddBiasColumn(currMat)
        End If

        currMat = MatMult(layerInput, T(W(layer)))
        
        ' now add transformations
        If layer = UBound(W) Then
            currMat = SigmoidMatrix(currMat)
        ElseIf Activations(layer) = "ReLu" Then
            currMat = ReLuMatrix(currMat)
        ElseIf Activations(layer) = "Sigmoid" Then
            currMat = SigmoidMatrix(currMat)
        End If
    Next layer
    
    ' evaluate loss and accuracy
    Dim testLoss As Double: testLoss = loss(currMat, testY)
    Dim testAcc As Double: testAcc = accuracy(currMat, testY)
    
    MsgBox "Test loss: " & testLoss & vbCrLf & "Test accuracy: " & testAcc
End Sub
