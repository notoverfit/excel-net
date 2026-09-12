Attribute VB_Name = "Utils"
Public Sub RaiseError(ByVal fn As String, ByVal msg As String)
    Err.Raise vbObjectError + 1000, fn, msg
End Sub

Public Sub PrintMatrix(ByVal mat As Variant)
    Dim r As Long, c As Long
    Dim rowText As String

    For r = 1 To UBound(mat, 1)
        rowText = ""

        For c = 1 To UBound(mat, 2)
            rowText = rowText & mat(r, c) & vbTab
        Next c

        Debug.Print rowText
    Next r
End Sub

Public Function loss(ByVal predictions As Variant, ByVal targets As Variant)
    Dim i As Long
    Dim lossSum As Double: lossSum = 0
    Dim lossCnt As Double: lossCnt = 0
    
    For i = LBound(predictions, 1) To UBound(predictions, 1)
        lossSum = lossSum + targets(i, 1) * Log(predictions(i, 1)) + (1 - targets(i, 1)) * Log(1 - predictions(i, 1))
        lossCnt = lossCnt + 1
    Next i
    
    loss = -lossSum / lossCnt
End Function

Public Function accuracy(ByVal predictions As Variant, ByVal targets As Variant)
    Dim correct As Long
    Dim total As Long
    Dim pred As Long
    
    For i = LBound(predictions, 1) To UBound(predictions, 1)
        If predictions(i, 1) > 0.5 Then
            pred = 1
        Else
            pred = 0
        End If
        
        If pred = targets(i, 1) Then
            correct = correct + 1
        End If
        total = total + 1
    Next i
    
    accuracy = correct / total
End Function

Public Sub PrintDimensions(mat As Variant)
    If Not IsArray(mat) Then
        Debug.Print "Not an array"
        Exit Sub
    End If
    
    Debug.Print "Rows: " & (UBound(mat, 1) - LBound(mat, 1) + 1)
    Debug.Print "Cols: " & (UBound(mat, 2) - LBound(mat, 2) + 1)
End Sub
