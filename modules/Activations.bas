Attribute VB_Name = "Activations"
Public Function dxReLu(ByVal x As Double) As Double
    If x > 0 Then
        dxReLu = 1
    Else
        dxReLu = 0
    End If
End Function

Public Function ReLu(ByVal x As Double) As Double
    If x > 0 Then
        ReLu = x
    Else
        ReLu = 0
    End If
End Function

Public Function dxSigmoid(ByVal x As Double) As Double
    dxSigmoid = Sigmoid(x) * (1 - Sigmoid(x))
End Function

Public Function Sigmoid(ByVal x As Double) As Double
    Sigmoid = 1 / (1 + Exp(-x))
End Function

Public Function ReLuMatrix(ByVal mat As Variant) As Variant
    Dim result() As Double
    Dim r As Long
    Dim c As Long
    
    ReDim result(1 To UBound(mat, 1), 1 To UBound(mat, 2))
    
    For r = 1 To UBound(mat, 1)
        For c = 1 To UBound(mat, 2)
            result(r, c) = ReLu(mat(r, c))
        Next c
    Next r
    
    ReLuMatrix = result
End Function

Public Function dxReLuMatrix(ByVal mat As Variant) As Variant
    Dim result() As Double
    Dim r As Long
    Dim c As Long
    
    ReDim result(1 To UBound(mat, 1), 1 To UBound(mat, 2))
    
    For r = 1 To UBound(mat, 1)
        For c = 1 To UBound(mat, 2)
            result(r, c) = dxReLu(mat(r, c))
        Next c
    Next r
    
    dxReLuMatrix = result
End Function

Public Function SigmoidMatrix(ByVal mat As Variant) As Variant
    Dim result() As Double
    Dim r As Long
    Dim c As Long
    
    ReDim result(1 To UBound(mat, 1), 1 To UBound(mat, 2))
    
    For r = 1 To UBound(mat, 1)
        For c = 1 To UBound(mat, 2)
            result(r, c) = Sigmoid(mat(r, c))
        Next c
    Next r
    
    SigmoidMatrix = result
End Function

Public Function dxSigmoidMatrix(ByVal mat As Variant) As Variant
    Dim result() As Double
    Dim r As Long
    Dim c As Long
    
    ReDim result(1 To UBound(mat, 1), 1 To UBound(mat, 2))
    
    For r = 1 To UBound(mat, 1)
        For c = 1 To UBound(mat, 2)
            result(r, c) = dxSigmoid(mat(r, c))
        Next c
    Next r
    
    dxSigmoidMatrix = result
End Function

