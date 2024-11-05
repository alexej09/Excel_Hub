Attribute VB_Name = "clean_cells"
Sub Fun_CleanUpCells(rng As Range)
    Dim cell As Range
    Dim strValue As String
    
    ' �berpr�fung, ob ein Bereich ausgew�hlt wurde
    If rng Is Nothing Then
        MsgBox "No range selected", vbExclamation
        Exit Sub
    End If
    
    ' Iteriere durch jede Zelle im ausgew�hlten Bereich
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
            
            ' Den bereinigten Wert zur�ck in die Zelle schreiben
            cell.Value = strValue
        End If
    Next cell
    
    MsgBox "Bereinigung abgeschlossen.", vbInformation
End Sub

Sub CleanUpCells()
    Dim rng As Range
    
    ' Benutzer w�hlt den Bereich aus
    On Error Resume Next
    Set rng = Application.InputBox("Bitte w�hlen Sie den Bereich aus, der �berpr�ft werden soll:", Type:=8)
    On Error GoTo 0
    
    ' Aufrufen der CleanUpRange-Routine
    Fun_CleanUpCells rng
End Sub
