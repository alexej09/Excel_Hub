'== constr_lookup ==
'Verwendung (Beispiele):
'  Vertikal:
'    =constr_lookup( (Suchspalte(A:A)=Suchwert(B1)*(Suchspalte(B:B)=Suchwert(C1))*(Suchspalte(N:N)=Suchwert(E1)) ; Outputspalte(D:D) )
'
' Analog f�r Suchzeilen und Werte!!!
'
'R�ckgabe:
'  � Standard: alle Treffer, dedupliziert, als String mit "; " getrennt
'  � F�r Spill-Ausgabe (jedes Ergebnis in eigener Zeile): drittes Argument auf "" setzen
'    z.B. =constr_lookup( (Suchspalte(A:A)=Suchwert(B1)*(Suchspalte(B:B)=Suchwert(C1))*(Suchspalte(N:N)=Suchwert(E1)) ; Outputspalte(D:D) ; "")
'
'Voraussetzungen:
'  � Die Kriterien-Ausdrucksmatrix (Produkt aus Bedingungen) muss die gleiche Ausrichtung/L�nge
'    wie die Output-Spalte/-Zeile haben (entweder 1 Spalte n Zeilen ODER 1 Zeile n Spalten).

Public Function constr_lookup(criteria As Variant, ret As Range, Optional sep As String = "; ") As Variant
    Dim arrCrit As Variant, vertical As Boolean, n As Long
    Dim found As New Collection, i As Long, critVal As Variant, retVal As Variant
    Dim out() As Variant, s As String

    'Kriteriums-Array aufnehmen (Range oder bereits berechnetes Array)
    If TypeName(criteria) = "Range" Then
        arrCrit = criteria.Value
    Else
        arrCrit = criteria
    End If

    'Ausrichtung der R�ckgabematrix ermitteln (einspaltig vs. einzeilig)
    If ret.Columns.Count = 1 And ret.Rows.Count >= 1 Then
        vertical = True
        n = ret.Rows.Count
    ElseIf ret.Rows.Count = 1 And ret.Columns.Count >= 1 Then
        vertical = False
        n = ret.Columns.Count
    Else
        constr_lookup = "#Output muss 1 Spalte ODER 1 Zeile sein"
        Exit Function
    End If

    'Dimensionen grob pr�fen
    On Error Resume Next
    If vertical Then
        If UBound(arrCrit, 1) < n Then
            constr_lookup = "#Kriterienl�nge passt nicht"
            Exit Function
        End If
    Else
        If UBound(arrCrit, 2) < n Then
            constr_lookup = "#Kriterienl�nge passt nicht"
            Exit Function
        End If
    End If
    On Error GoTo 0

    'Treffer sammeln (dedupliziert �ber Key)
    For i = 1 To n
        If vertical Then
            critVal = SafeIndex(arrCrit, i, 1)
            retVal = ret.Cells(i, 1).Value
        Else
            critVal = SafeIndex(arrCrit, 1, i)
            retVal = ret.Cells(1, i).Value
        End If

        If IsError(critVal) Or IsEmpty(critVal) Then
            'ignorieren
        Else
            If ToBool(critVal) Then
                On Error Resume Next
                found.Add retVal, CStr(retVal) 'dedupe
                On Error GoTo 0
            End If
        End If
    Next i

    If found.Count = 0 Then
        constr_lookup = "n.a."
        Exit Function
    End If

    'Ausgabe: Spill (sep="") oder zusammengefasst
    If sep = "" Then
        ReDim out(1 To found.Count, 1 To 1)
        For i = 1 To found.Count
            out(i, 1) = found(i)
        Next i
        constr_lookup = out          'spilled dynamic array
    Else
        For i = 1 To found.Count
            If i > 1 Then s = s & sep
            s = s & CStr(found(i))
        Next i
        constr_lookup = s
    End If
End Function

'--- Hilfsfunktionen ---
Private Function SafeIndex(a As Variant, r As Long, c As Long) As Variant
    On Error Resume Next
    SafeIndex = a(r, c)
    If Err.Number <> 0 Then
        'Einige Kriterien-Arrays kommen 1D daher (1..n)
        Err.Clear
        If r >= LBound(a) And r <= UBound(a) And (c = 1 Or c = 0) Then
            SafeIndex = a(r)
        End If
    End If
    On Error GoTo 0
End Function

Private Function ToBool(v As Variant) As Boolean
    'Akzeptiert TRUE/FALSE, 1/0, numerische Produkte der Bedingungen
    If VarType(v) = vbBoolean Then
        ToBool = v
    ElseIf IsNumeric(v) Then
        ToBool = (CDbl(v) <> 0)
    Else
        ToBool = (UCase$(CStr(v)) = "TRUE")
    End If
End Function

