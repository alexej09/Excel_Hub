Attribute VB_Name = "clean_word_tabs"
Sub clean_word_tabs()
    Dim BereichAuswahl As Range
    Dim NeueTabelle As Worksheet
    Dim TabellenName As String
    Dim Datum As String
    Dim Zelle As Range
    Dim MergedArea As Range
    Dim ZielZeile As Long
    Dim i As Long, k As Long, j As Long
    Dim Verkettet As String
    Dim ZeileHatVerbunden As Boolean
    Dim �berspringeBisZeile As Long

    ' Initialisieren
    ZielZeile = 1
    �berspringeBisZeile = 0

    ' Requirement 1: Bereich ausw�hlen
    On Error Resume Next
    Set BereichAuswahl = Application.InputBox("Bitte w�hlen Sie den Quellbereich aus:", "Bereich ausw�hlen", Type:=8)
    If BereichAuswahl Is Nothing Then
        MsgBox "Vorgang abgebrochen.", vbInformation
        Exit Sub
    End If
    If BereichAuswahl.Cells.Count = 0 Then
        MsgBox "Kein Bereich ausgew�hlt.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print "Bereich ausgew�hlt: " & BereichAuswahl.Address

    ' Requirement 2: Neue Tabelle erstellen
    Datum = Format(Date, "yyyymmdd")
    TabellenName = "�bertragung_" & Datum
    For Each NeueTabelle In ThisWorkbook.Worksheets
        If NeueTabelle.Name = TabellenName Then
            MsgBox "Eine Tabelle mit dem Namen '" & TabellenName & "' existiert bereits.", vbCritical
            Exit Sub
        End If
    Next NeueTabelle
    Set NeueTabelle = ThisWorkbook.Worksheets.Add
    NeueTabelle.Name = TabellenName

    Debug.Print "Neue Tabelle erstellt: " & NeueTabelle.Name

    For i = 1 To BereichAuswahl.Rows.Count
        ' �berspringe bereits verarbeitete Zeilen
        If i <= �berspringeBisZeile Then
            Debug.Print "Zeile " & i & " wird �bersprungen, da sie bereits verarbeitet wurde."
            ' Statt Continue For wird die Schleife durch Exit Sub oder Goto umgangen
            GoTo N�chsteZeile
        End If
    
        ZeileHatVerbunden = False ' Zur�cksetzen der Flagge f�r jede Zeile
    
        ' Durchlaufe jede Zelle der aktuellen Zeile
        For Each Zelle In BereichAuswahl.Rows(i).Cells
            If Zelle.MergeCells Then
                Set MergedArea = Zelle.MergeArea
    
                ' Nur die obere linke Zelle des Merge-Bereichs bearbeiten
                If Zelle.Address = MergedArea.Cells(1, 1).Address Then
                    Debug.Print "Verbundene Zelle erkannt bei: " & MergedArea.Address
    
                    ' Fall 2: Mehrere Spalten verbunden -> Abbruch
                    If MergedArea.Columns.Count > 1 Then
                        MsgBox "Mehrere Spalten sind verbunden. Verarbeitung wird abgebrochen.", vbCritical
                        Exit Sub
                    End If
    
                    ' Fall 1: Mehrere Zeilen verbunden -> Inhalte verarbeiten
                    If MergedArea.Rows.Count > 1 And MergedArea.Columns.Count = 1 Then
                        Debug.Print "Fall 1: Mehrere Zeilen verbunden erkannt bei: " & MergedArea.Address
                        ZeileHatVerbunden = True ' Flag setzen, da verbundene Zellen gefunden wurden
    
                        ' Verkettung der Inhalte aus benachbarten Spalten
                        For j = 1 To BereichAuswahl.Columns.Count
                            ' �berspringe die verbundene Spalte
                            If j = MergedArea.Column - BereichAuswahl.Column + 1 Then
                                ' Verbundene Zelle direkt �bertragen
                                NeueTabelle.Cells(ZielZeile, j).Value = MergedArea.Cells(1, 1).Value
                            Else
                                ' Inhalte der benachbarten Spalten verkettet �bertragen
                                Verkettet = ""
                                For k = 0 To MergedArea.Rows.Count - 1
                                    Verkettet = Verkettet & BereichAuswahl.Cells(MergedArea.Row + k, j).Value & Chr(10)
                                Next k
                                ' Entfernen des letzten Chr(10)
                                If Right(Verkettet, 1) = Chr(10) Then
                                    Verkettet = Left(Verkettet, Len(Verkettet) - 1)
                                End If
                                NeueTabelle.Cells(ZielZeile, j).Value = Verkettet
                                NeueTabelle.Cells(ZielZeile, j).WrapText = True
                            End If
                        Next j
    
                        ' Aktualisiere, bis zu welcher Zeile �bersprungen werden soll
                        �berspringeBisZeile = MergedArea.Row + MergedArea.Rows.Count - 1
                        ZielZeile = ZielZeile + 1
                        GoTo N�chsteZeile ' Gehe zur n�chsten Zeile, da die aktuelle abgeschlossen ist
                    End If
                End If
            End If
        Next Zelle
    
        ' Fall: Keine verbundenen Zellen in der Zeile
        If Not ZeileHatVerbunden Then
            Debug.Print "Keine verbundenen Zellen in Zeile " & i & ". Daten werden eins zu eins �bertragen."
            For j = 1 To BereichAuswahl.Columns.Count
                NeueTabelle.Cells(ZielZeile, j).Value = BereichAuswahl.Cells(i, j).Value
            Next j
            ZielZeile = ZielZeile + 1
        End If
    
N�chsteZeile:
    Next i
    ' Abschluss: Gesamte Tabelle oben ausrichten
    NeueTabelle.Cells.VerticalAlignment = xlTop
    ' Abschlussmeldung
    MsgBox "Verarbeitung abgeschlossen.", vbInformation
End Sub

