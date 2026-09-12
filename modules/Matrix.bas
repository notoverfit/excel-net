Attribute VB_Name = "Matrix"
Public Function RemoveBiasColumn(mat As Variant) As Variant
    Dim rows As Long
    Dim cols As Long
    Dim result() As Double
    Dim r As Long
    Dim c As Long

    rows = UBound(mat, 1)
    cols = UBound(mat, 2)

    ReDim result(1 To rows, 1 To cols - 1)

    For r = 1 To rows
        For c = 1 To cols - 1
            result(r, c) = mat(r, c)
        Next c
    Next r

    RemoveBiasColumn = result
End Function

Public Function AddBiasColumn(mat As Variant) As Variant
    Dim rows As Long
    Dim cols As Long
    Dim result() As Double
    Dim r As Long, c As Long
    
    rows = UBound(mat, 1)
    cols = UBound(mat, 2)
    
    ReDim result(1 To rows, 1 To cols + 1)
    
    For r = 1 To rows
        For c = 1 To cols
            result(r, c) = mat(r, c)
        Next c
        
        ' final column = bias input
        result(r, cols + 1) = 1#
    Next r
    
    AddBiasColumn = result
End Function

Public Function MatMult(ByVal A As Variant, ByVal B As Variant) As Variant

    Dim result() As Double
    Dim i As Long, j As Long, k As Long
    
    Dim rowsA As Long
    Dim colsA As Long
    Dim rowsB As Long
    Dim colsB As Long
    
    rowsA = UBound(A, 1)
    colsA = UBound(A, 2)
    rowsB = UBound(B, 1)
    colsB = UBound(B, 2)
    
    If colsA <> rowsB Then
        Call RaiseError("MatMult", "Inner matrix dimensions do not match")
    End If
    
    ReDim result(1 To rowsA, 1 To colsB)
    
    For i = 1 To rowsA
        For j = 1 To colsB
            For k = 1 To colsA
                result(i, j) = result(i, j) + A(i, k) * B(k, j)
            Next k
        Next j
    Next i
    
    MatMult = result

End Function

Public Function T(ByVal mat As Variant) As Variant
    Dim result() As Double
    Dim r As Long, c As Long
    Dim rows As Long, cols As Long

    rows = UBound(mat, 1)
    cols = UBound(mat, 2)

    ReDim result(1 To cols, 1 To rows)

    For r = 1 To rows
        For c = 1 To cols
            result(c, r) = mat(r, c)
        Next c
    Next r

    T = result
End Function

Public Function MatSubtract(ByVal m1 As Variant, ByVal m2 As Variant) As Variant
    Dim result() As Double
    Dim r As Long, c As Long

    If UBound(m1, 1) <> UBound(m2, 1) Or UBound(m1, 2) <> UBound(m2, 2) Then
        Call RaiseError("MatSubtract", "Dimensions of matrices do not match")
    End If

    ReDim result(1 To UBound(m1, 1), 1 To UBound(m1, 2))

    For r = 1 To UBound(m1, 1)
        For c = 1 To UBound(m1, 2)
            result(r, c) = m1(r, c) - m2(r, c)
        Next c
    Next r

    MatSubtract = result
End Function

Function ElementMultiply(mat As Variant, scalar As Double) As Variant
    Dim rows As Long
    Dim cols As Long
    Dim r As Long
    Dim c As Long
    
    rows = UBound(mat, 1)
    cols = UBound(mat, 2)
    
    Dim result() As Double
    ReDim result(1 To rows, 1 To cols)
    
    For r = 1 To rows
        For c = 1 To cols
            result(r, c) = mat(r, c) * scalar
        Next c
    Next r
    
    ElementMultiply = result
End Function

Public Function MatrixElementMultiply(ByVal m1 As Variant, _
                                      ByVal m2 As Variant) As Variant
    Dim result() As Double
    Dim r As Long, c As Long

    If UBound(m1, 1) <> UBound(m2, 1) Or _
       UBound(m1, 2) <> UBound(m2, 2) Then
        Call RaiseError("MatrixElementMultiply", _
                        "Matrix dimensions do not match")
    End If

    ReDim result(1 To UBound(m1, 1), 1 To UBound(m1, 2))

    For r = 1 To UBound(m1, 1)
        For c = 1 To UBound(m1, 2)
            result(r, c) = m1(r, c) * m2(r, c)
        Next c
    Next r

    MatrixElementMultiply = result
End Function

Public Sub PrintDims(ByVal mat As Variant)
    Debug.Print name & ": " & UBound(mat, 1) & " x " & UBound(mat, 2)
End Sub

