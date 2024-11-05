Attribute VB_Name = "SepString_in_rows"
Sub Fun_CleanUpCells(rng As Range)
    Dim cell As Range
    Dim strValue As String
    
    ' Überprüfung, ob ein Bereich ausgewählt wurde
    If rng Is Nothing Then
        MsgBox "No range selected", vbExclamation
        Exit Sub
    End If
    
    ' Iteriere durch jede Zelle im ausgewählten Bereich
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            strValue = cell.Value
            
            ' Entfernen von Chr(10), Chr(13) und Leerzeichen am Anfang und Ende
            Do While Left(strValue, 1) = Chr(10) Or Left(strValue, 1) = Chr(13) Or Left(strValue, 1) = " "
                strValue = Mid(strValue, 2)
            Loop
            
            Do While Right(strValue, 1) = Chr(10) Or Right(strValue, 1) = Chr(13) Or Right(strValue, 1) = " "
                strValue = Left(strValue, Len(strValue) - 1)
            Loop
            
            ' Den bereinigten Wert zurück in die Zelle schreiben
            cell.Value = strValue
        End If
    Next cell
    
    MsgBox "Bereinigung abgeschlossen.", vbInformation
End Sub

Sub SepString_in_rows()

    Dim St As Long ' Variable used for start of a string
    Dim Del As Long ' Variable used for string length
    Dim Sep As New Collection ' List of place numbers of the signs used for separation
    Dim Lst As New Collection ' List with the separated statements
    Dim RowNr As New Collection ' List of Row numbers corresponding to found strings

    On Error Resume Next

' Select Column with the Text which need to be separated
    Dim Rn As Range
        Set Rn = Application.InputBox( _
            Title:="Separate Strings in Rows", _
            Prompt:="Select a column with Text", _
            Type:=8)
'    On Error GoTo LabelExit

    Fun_CleanUpCells Rn
    
    Debug.Print "Range = " & Rn.Address
    Dim n As String ' Separation sign
    n = Chr(10)

    Dim LL As Long ' Length of the selected range
    LL = Rn.Rows.Count
'    Debug.Print "LL= " & LL
'    Define List of entries which separate the statements

    Dim L As Long ' Length of the string in a Cell
    
    For s = 1 To LL
        L = Len(Rn.Cells(s, 1).Value)
'        Debug.Print "L= " & L
'        Debug.Print "s= " & s
        
        For i = 1 To L

            If Mid(Rn.Cells(s, 1).Value, i, 1) = n Then
            
                Sep.Add (i)

            ElseIf i = L Then

                Sep.Add (L + 1)
                                
            End If
        
        Next i
    
        Dim m As Long ' number of signs defined for separation
        
        If Right(Rn.Cells(s, 1).Value, 1) = n Then ' If the last sign is an Enter take it into the list
        
            m = Sep.Count + 1
            
        Else
        
            m = Sep.Count
            
        End If
'        Debug.Print "m= " & m
        
    '    Define List of entries which need to be checked for duplicates

        For j = 1 To m
        RowNr.Add Rn.Cells(s, 1).Row
     
'        Debug.Print "j= " & j
        
            If j = 1 Then
                   
                Lst.Add (Left(Rn.Cells(s, 1).Value, Sep(j) - 1))
    
            ElseIf j = m Then
    
                Lst.Add (Right(Rn.Cells(s, 1).Value, L - Sep(m - 1)))
                
            Else
            
                St = Sep(j - 1) + 1
                Del = Sep(j) - Sep(j - 1) - 1
                Lst.Add Mid(Rn.Cells(s, 1).Value, St, Del)
                
            End If
'        Debug.Print "Sep(j)= " & Sep(j)
'        Debug.Print "Lst(j)= " & Lst(j)
        
        Next j
        
        Set Sep = Nothing

    Next s
    
    '    Separate the statements in Rows
    
        ss = Sep.Count
        rr = RowNr.Count
        kk = Lst.Count
        
'        Debug.Print "Number of Enter = " & ss
'        Debug.Print "Number of Rows = " & rr
'        Debug.Print "Number of Statements = " & kk
        
    '    Create new Worksheet where the separated strings will be stored
        Set newwks = Worksheets.Add
        
        With newwks
            .Name = "Sep_Strings " & Format(Now(), "hh-nn-ss")
            .ListObjects.Add(xlSrcRange, Range("A1:C1"), , xlNo).Name = "Strings " & Format(Now(), "hh-nn-ss")
            .Range("A1:C1").Value = Array("Row Number", "Separated Strings", "Duplicate Check")
            .Columns("B:B").ColumnWidth = 90
            .Columns("B:B").WrapText = True
        End With
        
        For k = 1 To kk
        
            newwks.Cells(k + 1, 1).Value = RowNr(k)
            newwks.Cells(k + 1, 2).Value = Lst(k)
            newwks.Cells(k + 1, 3).Formula2R1C1 = "=COUNTIFS([Row Number],[@[Row Number]],[Separated Strings],[@[Separated Strings]])"
        
        Next k

LabelExit:


End Sub

