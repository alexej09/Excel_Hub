Public Function suplookup(myStr As String, srch As Range, ret As Range) As String
    Dim output As String
    Dim j As Integer
    Dim i As Integer
    Dim myRow As Integer
    Dim myCol As Integer
    Dim srchColAd As Integer
    Dim retColAd As Integer
    Dim srchRowAd As String
    Dim retRowAd As String
    Dim foundValues As New Collection

    myRow = srch.Rows.Count
    myCol = srch.Columns.Count
    srchColAd = Val(Right(srch.Cells(1, 1).Address, 1))
    retColAd = Val(Right(ret.Cells(1, 1).Address, 1))
    srchRowAd = Left(srch.Cells(1, 1).Address, 2)
    retRowAd = Left(ret.Cells(1, 1).Address, 2)
    Debug.Print retRowAd

    If myRow > myCol And srchColAd = retColAd Then
        For i = 1 To myRow
            If InStr(srch(i), myStr) > 0 Then
                On Error Resume Next
                foundValues.Add ret(i), CStr(ret(i))
                On Error GoTo 0
            End If
        Next i
    ElseIf myRow < myCol And srchRowAd = retRowAd Then
        For i = 1 To myCol
            If InStr(srch(i), myStr) > 0 Then
                On Error Resume Next
                foundValues.Add ret(i), CStr(ret(i))
                On Error GoTo 0
            End If
        Next i
    Else
        suplookup = "#Check selected Ranges"
        Exit Function
    End If

    If foundValues.Count > 0 Then
        For Each item In foundValues
            output = output & "; " & item
        Next item
        ' Entfernt das erste "; " für eine saubere Ausgabe:
        suplookup = Mid(output, 3)
    Else
        suplookup = "n.a."
    End If
End Function


