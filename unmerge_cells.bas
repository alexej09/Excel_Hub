Attribute VB_Name = "unmerge_cells"
Sub Unmerge()
Dim Rn As Range

' Request to select a range
    On Error Resume Next
        Set Rn = Application.InputBox( _
            Title:="Unmerge Cells in Range", _
            Prompt:="Select Range for application", _
            Type:=8)
            
' Check Range for Merged Cells
    For Each Cel In Rn.Cells
        If Cel.MergeCells And Cel.MergeArea.Rows.Count > 1 Then
        
' Check if Merged Cells are vertikal(repeate values for each Row)
            With Cel.MergeArea
                .Unmerge
                .Value = Cel.Value
            End With
            
' Check if Merged Cells are purly horizontal (do not repeate the values)
        ElseIf Cel.MergeCells Then
        
            With Cel.MergeArea
                .Unmerge
                .HorizontalAlignment = xlCenterAcrossSelection
            End With
            
        End If
    Next Cel
    
        a = Rn.Rows.Count
        b = Rn.Columns.Count


                For j = 1 To b
                    For i = 2 To a

                            ' check of values and fontcolor of the first column
                            If j = 1 And Rn.Cells(i - 1, 1).Value = Rn.Cells(i, 1).Value And IsEmpty(Rn.Cells(i, 1).Value) = False Then
                            
                                Rn.Cells(i, 1).Font.ColorIndex = 15
'                                Debug.Print "j= " & j
'                                Debug.Print "Zelle oben " & Rn.Cells(i - 1, 1).Value
'                                Debug.Print "Zelle " & Rn.Cells(i, 1).Value
                            End If

                            
                            ' check of values and fontcolor of the remained colums
'                           If j > 1 And Rn.Cells(i - 1, j).Value = Rn.Cells(i, j).Value And Rn.Cells(i, j - 1).Font.ColorIndex = 15 And IsEmpty(Rn.Cells(i, j).Value) = False Then
                            If j > 1 Then
                                If Rn.Cells(i - 1, j).Value = Rn.Cells(i, j).Value And IsEmpty(Rn.Cells(i, j).Value) = False And Rn.Cells(i, j - 1).Font.ColorIndex = 15 Then
                            
                                Rn.Cells(i, j).Font.ColorIndex = 15

'                                Debug.Print "Color " & Rn.Cells(i, j - 1).Font.ColorIndex
                                End If
                                
                            End If


                        ' check top boarder
                        If Rn.Cells(i, j).Font.ColorIndex = 15 And IsEmpty(Rn.Cells(i, j).Value) = False Then
                        Rn.Cells(i, j).Borders(xlEdgeTop) = none

                        End If


                Next i
            Next j
       
    On Error GoTo 0
   
End Sub
