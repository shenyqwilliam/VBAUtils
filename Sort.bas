Attribute VB_Name = "Sort"
Option Explicit



Sub BENCHMARK()
    Dim sampleIndex$: sampleIndex = "_rnd"
    Dim arrIO: arrIO = Sheets("sample").Range(sampleIndex).Value
    Dim arrProc As Variant, arrRef As Variant, i&
    Dim errPos&
    
    'Convert 2-dimension array to 1-dimension
    ReDim arrProc(LBound(arrIO, 1) To UBound(arrIO, 1))
    For i = LBound(arrIO, 1) To UBound(arrIO, 1)
        arrProc(i) = arrIO(i, 1)
    Next i
    
    
    '*** Call the sort algorithm ***
    arrRef = QuickSort(arrProc)
    
    
    'Verify the result
    For i = LBound(arrProc) To UBound(arrProc) - 1
        If arrProc(i) > arrProc(i + 1) Then
            errPos = i
            MsgBox "Failed! arr(" & i & ")=" & arrProc(i) & " > arr(" & i + 1 & ")=" & arrProc(i + 1)
            Exit For
        End If
    Next i
    
    
    'Convert 1-dimension array to 2-dimension and Output to sheet
    For i = LBound(arrIO, 1) To UBound(arrIO, 1)
        arrIO(i, 1) = arrProc(i)
    Next i
    Sheets("sample").Range("_out").Value = arrIO
    
    For i = LBound(arrIO, 1) To UBound(arrIO, 1)
        arrIO(i, 1) = arrRef(i)
    Next i
    Sheets("sample").Range("_ref").Value = arrIO
End Sub



Public Function QuickSort(ByRef arr As Variant) As Long()
    Dim i&
    Dim lo&, hi&
    Dim arr_ref&()
    
    lo = LBound(arr): hi = UBound(arr)
    
    ReDim arr_ref(lo To hi)
    For i = LBound(arr_ref) To UBound(arr_ref): arr_ref(i) = i: Next i
    
    Call qsort(arr:=arr, lo:=lo, hi:=hi, arr_ref:=arr_ref)
    
    QuickSort = arr_ref
End Function



Private Sub qsort(ByRef arr As Variant, ByVal lo&, ByVal hi&, ByRef arr_ref As Variant)
    Dim md&     'middle index
    Dim p#      'pivot
    Dim s#      'temp variable for swapping
    Dim m&, n&  'paired indices for partitioning
    
    
    'Choosing the pivot by "Median-of-three"
    'Note the swap here is a must, because we don't swap the element at end if all elements are less/greater than the pivot
    md = lo + (hi - lo) \ 2
    If arr(lo) > arr(hi) Then   'swap lo<=>hi
        s = arr(lo): arr(lo) = arr(hi): arr(hi) = s
        s = arr_ref(lo): arr_ref(lo) = arr_ref(hi): arr_ref(hi) = s
    End If
    If arr(lo) > arr(md) Then   'swap lo<=>md
        s = arr(lo): arr(lo) = arr(md): arr(md) = s
        s = arr_ref(lo): arr_ref(lo) = arr_ref(md): arr_ref(md) = s
    End If
    If arr(md) > arr(hi) Then   'swap md<=>hi
        s = arr(md): arr(md) = arr(hi): arr(hi) = s
        s = arr_ref(md): arr_ref(md) = arr_ref(hi): arr_ref(hi) = s
    End If
    p = arr(md)
    
    
    'Set paired indices at two ends
    'Note we already sorted lo and hi, so we can skip them
    m = lo + 1: n = hi - 1
    
    'Squeeze the paired indices
    'Exchange the paired value whenever they are on both sides of pivot value
    'Loop till they cross or meet at either end
    Do While (m < n)
        Do While m < hi And arr(m) <= p
            m = m + 1
        Loop
        
        Do While n > lo And arr(n) >= p
            n = n - 1
        Loop
        
        If m < n Then   'swap m<=>n
            s = arr(m): arr(m) = arr(n): arr(n) = s
            s = arr_ref(m): arr_ref(m) = arr_ref(n): arr_ref(n) = s
            m = m + 1: n = n - 1
        ElseIf m = n Then
            m = m + 1: n = n - 1
        End If
    Loop
    
    
    'Sort subsets
    If n > lo Then Call qsort(arr, lo, n, arr_ref)
    If m < hi Then Call qsort(arr, m, hi, arr_ref)
End Sub

