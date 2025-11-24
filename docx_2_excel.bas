Attribute VB_Name = "Modul5"
Option Explicit

' === Einstieg: Dateiauswahl für EIN Word-Dokument ===
Public Sub ImportWord_FromPicker_UI()
    Dim fd As Object, selPath As String
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .Title = "Word-Dokument auswählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word-Dokumente", "*.doc;*.docx"
        If Len(ThisWorkbook.Path) > 0 Then .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        If .Show <> -1 Then Exit Sub
        selPath = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ProcessWordFile selPath
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' --- Word öffnen, Tabellen auslesen; Merge-Zellen: Werte über Bereich wiederholen ---
Private Sub ProcessWordFile(ByVal filePath As String)
    Dim wdApp As Object, wdDoc As Object, wdTbl As Object
    Dim t As Long, r As Long, c As Long, rowsCt As Long, colsCt As Long
    Dim ws As Worksheet, docBase As String
    Dim arr() As Variant, have() As Boolean
    Dim v As Variant, ok As Boolean
    Dim lastLeft As Variant
    
    On Error GoTo EH
    
    docBase = GetFileBaseName(filePath)
    Set wdApp = GetOrCreateWordApp()
    If wdApp Is Nothing Then Exit Sub
    
    Set wdDoc = wdApp.Documents.Open(Filename:=filePath, ReadOnly:=True, Visible:=False)
    If wdDoc.Tables.Count = 0 Then GoTo QuitDoc
    
    For t = 1 To wdDoc.Tables.Count
        Set wdTbl = wdDoc.Tables(t)
        rowsCt = wdTbl.Rows.Count
        colsCt = wdTbl.Columns.Count
        
        ReDim arr(1 To rowsCt, 1 To colsCt)
        ReDim have(1 To rowsCt, 1 To colsCt)
        
        ' 1) Rohwerte lesen; fehlende (wegen Merge) bleiben False
        For r = 1 To rowsCt
            For c = 1 To colsCt
                v = GetCellTextSafe(wdTbl, r, c, ok)
                If ok Then
                    arr(r, c) = v
                    have(r, c) = True
                End If
            Next c
        Next r
        
        ' 2) Links füllen (horizontale Merges replizieren)
        For r = 1 To rowsCt
            lastLeft = vbNullString
            For c = 1 To colsCt
                If have(r, c) And Len(arr(r, c)) > 0 Then
                    lastLeft = arr(r, c)
                ElseIf Not have(r, c) And Len(lastLeft) > 0 Then
                    arr(r, c) = lastLeft
                    have(r, c) = True
                End If
            Next c
        Next r
        
        ' 3) Von oben füllen (vertikale Merges replizieren)
        For r = 2 To rowsCt
            For c = 1 To colsCt
                If Not have(r, c) And have(r - 1, c) Then
                    arr(r, c) = arr(r - 1, c)
                    have(r, c) = True
                End If
            Next c
        Next r
        
        ' 4) Neues Blatt + Array in einem Rutsch schreiben
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        On Error Resume Next
        ws.Name = SanitizeSheetName(Left$(docBase & " - Tabelle " & CStr(t), 31))
        Err.Clear: On Error GoTo EH
        ws.Range(ws.Cells(1, 1), ws.Cells(rowsCt, colsCt)).Value = arr
        
        ' 5) In intelligente Tabelle (erste Zeile = Überschriften) und AutoFit
        Dim lo As ListObject
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range(ws.Cells(1, 1), ws.Cells(rowsCt, colsCt)), XlListObjectHasHeaders:=xlYes)
        ws.Columns.AutoFit
    Next t
    
QuitDoc:
    wdDoc.Close SaveChanges:=False
    Exit Sub

EH:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    MsgBox "Fehler bei Datei: " & filePath & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' --- Sicherer Zugriff auf Word-Zelle: gibt ok=False zurück, wenn Position durch Merge fehlt ---
Private Function GetCellTextSafe(ByVal wdTbl As Object, ByVal r As Long, ByVal c As Long, ByRef ok As Boolean) As String
    Dim txt As String
    On Error Resume Next
    txt = wdTbl.cell(r, c).Range.Text
    If Err.Number <> 0 Then
        ok = False
        Err.Clear
        Exit Function
    End If
    ok = True
    GetCellTextSafe = CleanCellText(txt)
End Function

' --- Word-Instanz (late binding) ---
Private Function GetOrCreateWordApp() As Object
    On Error Resume Next
    Set GetOrCreateWordApp = GetObject(, "Word.Application")
    If GetOrCreateWordApp Is Nothing Then
        Set GetOrCreateWordApp = CreateObject("Word.Application")
        If Not GetOrCreateWordApp Is Nothing Then GetOrCreateWordApp.Visible = False
    End If
    On Error GoTo 0
End Function

' --- Hilfen ---
Private Function GetFileBaseName(ByVal filePath As String) As String
    Dim p As Long
    p = InStrRev(filePath, Application.PathSeparator)
    If p > 0 Then
        GetFileBaseName = Mid$(filePath, p + 1)
    Else
        GetFileBaseName = filePath
    End If
    If InStrRev(GetFileBaseName, ".") > 0 Then
        GetFileBaseName = Left$(GetFileBaseName, InStrRev(GetFileBaseName, ".") - 1)
    End If
End Function

Private Function CleanCellText(ByVal s As String) As String
    Dim tmp As String
    tmp = Replace$(s, Chr$(13) & Chr$(7), vbNullString) ' End-of-cell
    tmp = Replace$(tmp, Chr$(7), vbNullString)
    tmp = Replace$(tmp, vbCr, vbLf)
    CleanCellText = Trim$(tmp)
End Function

Private Function SanitizeSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), " ")
    Next i
    If Len(Trim$(s)) = 0 Then s = "Tabelle"
    SanitizeSheetName = s
End Function

