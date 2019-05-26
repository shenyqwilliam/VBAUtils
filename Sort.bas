Attribute VB_Name = "Sort"
Sub BENCHMARK()
    Dim sampleIndex$: sampleIndex = "_rnd"
    Dim arrIO: arrIO = Sheets("sample").Range(sampleIndex).Value
    Dim arrProc, i&
    Dim errPos&
    
    'Convert 2-dimension array to 1-dimension
    ReDim arrProc(LBound(arrIO, 1) To UBound(arrIO, 1))
    For i = LBound(arrIO, 1) To UBound(arrIO, 1)
        arrProc(i) = arrIO(i, 1)
    Next i
    
    
    '*** Call the sort algorithm ***
    Call QuickSort(arrProc, LBound(arrProc), UBound(arrProc))
    
    
    'Verify the result
    For i = LBound(arrProc) To UBound(arrProc) - 1
        If arrProc(i) > arrProc(i + 1) Then
            errPos = i
            MsgBox "Failed! arr(" & i & ")=" & arrProc(i) & " > arr(" & i + 1 & ")=" & arrProc(i + 1)
            Exit For
        End If
    Next i
    
    'Convert 1-dimension array to 2-dimension
    For i = LBound(arrIO, 1) To UBound(arrIO, 1)
        arrIO(i, 1) = arrProc(i)
    Next i
    Sheets("sample").Range("_output").Value = arrIO
End Sub


'===============================================

Public Sub QuickSort(ByRef arr As Variant, ByVal lo&, ByVal hi&, Optional ByVal preScan As Boolean = False)
    Dim md&     'mid
    Dim p#      'pivot
    Dim s#      'temp variable for swapping
    Dim m&, n&  'paired indices for partitioning
    Dim i&, isSorted As Boolean
    
    'Pre-run
    If preScan = True Then
        isSorted = True
        For i = lo To hi - 1
            If arr(i) > arr(i + 1) Then isSorted = False: Exit For
        Next i
        If isSorted = True Then Exit Sub
    End If
    
    'Choosing the pivot by "Median-of-three"
    'Note the swap here is a must because we don't swap an element equal to pivot
    md = lo + (hi - lo) \ 2
    If arr(lo) > arr(hi) Then s = arr(lo): arr(lo) = arr(hi): arr(hi) = s       'swap lo<=>hi
    If arr(lo) > arr(md) Then s = arr(lo): arr(lo) = arr(md): arr(md) = s       'swap lo<=>md
    If arr(md) > arr(hi) Then s = arr(md): arr(md) = arr(hi): arr(hi) = s       'swap md<=>hi
    p = arr(md)
    
    'Set paired indices at two ends
    'Note we already sorted lo and hi, so we can skip them
    m = lo + 1: n = hi - 1
    
    Do While (m < n)
        Do While m < hi And arr(m) <= p
            m = m + 1
        Loop
        
        Do While n > lo And arr(n) >= p
            n = n - 1
        Loop
        
        If m < n Then
            s = arr(m): arr(m) = arr(n): arr(n) = s
            m = m + 1: n = n - 1
        ElseIf m = n Then
            m = m + 1
            n = n - 1
        End If
    Loop
    
    If n > lo Then Call QuickSort(arr, lo, n, preScan)
    If m < hi Then Call QuickSort(arr, m, hi, preScan)
End Sub

