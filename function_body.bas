Option Explicit

Function Hrz_UNIQUE(Source_range As Range, index As Long) As Variant
    Dim value_i As Variant
    Dim unique_count As Long, int_i As Long
    Dim value_list() As Variant, unique_list() As Variant
    ReDim unique_list(1 To 1)
    value_list = Source_range.value
    unique_count = 1
    For Each value_i In value_list
        If IsError(Application.Match(value_i, unique_list, 0)) Then
            ReDim Preserve unique_list(1 To unique_count)
            unique_list(unique_count) = value_i
            unique_count = unique_count + 1
        End If
    Next value_i
    On Error GoTo ErrorNA
        Hrz_UNIQUE = unique_list(index)
    On Error GoTo 0
    Exit Function
ErrorNA:
    Hrz_UNIQUE = CVErr(xlErrNA)
    Exit Function
End Function

Function Hrz_SORTBY(Value_range As Range, Name_range As Range, index As Long, mode As Boolean) As Variant
    Dim value_list() As Variant, name_list() As Variant
    value_list = Value_range.value
    name_list = Name_range.value
    If mode = False Then
        Call QuickSortDescending2(value_list, name_list, LBound(value_list), UBound(value_list))
    Else
        Call QuickSortAscending2(value_list, name_list, LBound(value_list), UBound(value_list))
    End If
    On Error GoTo ErrorNA
        Hrz_SORTBY = name_list(index, 1)
    On Error GoTo 0
    Exit Function
ErrorNA:
    Hrz_SORTBY = CVErr(xlErrNA)
    Exit Function
End Function

Sub QuickSortDescending2(ByRef value_list() As Variant, ByRef name_list() As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant
    i = low
    j = high
    pivot = value_list((low + high) \ 2, 1)
    Do While i <= j
        Do While value_list(i, 1) > pivot
            i = i + 1
        Loop
        Do While value_list(j, 1) < pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = value_list(i, 1)
            value_list(i, 1) = value_list(j, 1)
            value_list(j, 1) = temp
            temp = name_list(i, 1)
            name_list(i, 1) = name_list(j, 1)
            name_list(j, 1) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then Call QuickSortDescending2(value_list, name_list, low, j)
    If i < high Then Call QuickSortDescending2(value_list, name_list, i, high)
End Sub

Sub QuickSortAscending2(ByRef value_list() As Variant, ByRef name_list() As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant
    i = low
    j = high
    pivot = value_list((low + high) \ 2, 1)
    Do While i <= j
        Do While value_list(i, 1) < pivot
            i = i + 1
        Loop
        Do While value_list(j, 1) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = value_list(i, 1)
            value_list(i, 1) = value_list(j, 1)
            value_list(j, 1) = temp
            temp = name_list(i, 1)
            name_list(i, 1) = name_list(j, 1)
            name_list(j, 1) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then Call QuickSortAscending2(value_list, name_list, low, j)
    If i < high Then Call QuickSortAscending2(value_list, name_list, i, high)
End Sub

Function Hrz_SORTBY1(Value_range As Range, Name_range As Range, index As Long, mode As Boolean) As Variant
    Dim value_list() As Variant, name_list() As Variant, value_list1() As Variant, name_list1() As Variant
    value_list = Value_range.value
    name_list = Name_range.value
    Dim i As Long, j As Long
    For i = UBound(value_list) To 1 Step -1
        If value_list(i, 1) <> 0 Then
            Exit For
        End If
    Next i
    ReDim Preserve value_list1(1 To i)
    ReDim Preserve name_list1(1 To i)
    For j = 1 To i
        value_list1(j) = value_list(j, 1)
        name_list1(j) = name_list(j, 1)
    Next j
    If mode = False Then
        Call QuickSortDescending1(value_list1, name_list1, LBound(value_list1), UBound(value_list1))
    Else
        Call QuickSortAscending1(value_list1, name_list1, LBound(value_list1), UBound(value_list1))
    End If
    On Error GoTo ErrorNA
        Hrz_SORTBY1 = name_list1(index)
    On Error GoTo 0
    Exit Function
ErrorNA:
    Hrz_SORTBY1 = CVErr(xlErrNA)
    Exit Function
End Function

Sub QuickSortDescending1(ByRef value_list() As Variant, ByRef name_list() As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant
    i = low
    j = high
    pivot = value_list((low + high) \ 2)
    Do While i <= j
        Do While value_list(i) > pivot
            i = i + 1
        Loop
        Do While value_list(j) < pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = value_list(i)
            value_list(i) = value_list(j)
            value_list(j) = temp
            temp = name_list(i)
            name_list(i) = name_list(j)
            name_list(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then Call QuickSortDescending1(value_list, name_list, low, j)
    If i < high Then Call QuickSortDescending1(value_list, name_list, i, high)
End Sub

Sub QuickSortAscending1(ByRef value_list() As Variant, ByRef name_list() As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant
    i = low
    j = high
    pivot = value_list((low + high) \ 2)
    Do While i <= j
        Do While value_list(i) < pivot
            i = i + 1
        Loop
        Do While value_list(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = value_list(i)
            value_list(i) = value_list(j)
            value_list(j) = temp
            temp = name_list(i)
            name_list(i) = name_list(j)
            name_list(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then Call QuickSortAscending1(value_list, name_list, low, j)
    If i < high Then Call QuickSortAscending1(value_list, name_list, i, high)
End Sub
