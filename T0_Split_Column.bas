Attribute VB_Name = "T0_Split_Column"
Sub SplitColumnBezeichnung()
    Dim asSheet As Worksheet, wsA As Worksheet, wsB As Worksheet, wsC As Worksheet, wsD As Worksheet, wsE As Worksheet, wsF As Worksheet, wsG As Worksheet, wsH As Worksheet, wsI As Worksheet
    Dim headerRange As Variant
    Dim LOwstA As ListObject
    Dim headersA As Range, dataRangeA As Range
    Dim match_array() As Variant
    Dim colIndexB As Integer, lastrow As Integer, lastColumnA As Integer, y As Integer
    Dim colNameB As String, columnlematcherB As String, sourceString As String
    
    ' Set references to worksheets
    Set wstA = ThisWorkbook.Sheets("Page029")
    Set asSheet = ThisWorkbook.Sheets("Artikelstamm")
    Set LOwstA = wstA.ListObjects("Page029")
    Set headersA = wstA.Range("A1").Resize(1, wstA.Cells(1, wstA.Columns.Count).End(xlToLeft).Column)
    
    ' Get the string array from the specified range in "Artikelstamm" worksheet
    stringArray = asSheet.Range("A1:A" & asSheet.Cells(asSheet.Rows.Count, "A").End(xlUp).Row).Value
    ' Debug.Print "Artikelstamm: " & stringArray(10, 1)
    
    ' Get lastcolumn and lastrow from wstA
    lastRowA = wstA.Cells(wstA.Rows.Count, "B").End(xlUp).Row
    lastColumnA = wstA.Cells(1, wstA.Columns.Count).End(xlToLeft).Column
        
    ' Set header name And get column index defined by header name
    colNameB = "Bezeichnung"
    colIndexB = getColIndex(headersA, colNameB, lastColumnA)
    columnlematcherB = wstA.Cells(1, colIndexB).Address
    
    ' Define header range
    Set dataRangeA = wstA.Range(columnlematcherB).Resize(lastRowA - 1, headersA.Columns.Count)
        
    ' Paste new Column Masse if not exist
    Dim check As Boolean
    check = checkcolumnMasse(headersA)
    'Debug.Print "checkcolumnMasse: " & check
    If check <> True Then
        LOwstA.ListColumns.Add(colIndexB + 1).Name = "Masse"
    End If
    
    ' Get array of matches in Column Bezeichnung with Artikelstamm
    ReDim match_array(1 To lastRowA, 1 To 2) As Variant
    
    y = 1
    ' Loop through column 'Bezeichnung' from table wstA '
    For x = 2 To lastRowA
        ' Get Values from each cell
        sourceString = wstA.Cells(x, colIndexB)
        For j = LBound(stringArray, 1) To UBound(stringArray, 1)
            ' Check Cell Value in Column Bezeichnung if there is a match in the 'Artikelstamm' sheet'
            If InStr(1, sourceString, stringArray(j, 1)) > 0 Then
            ' If there is a match add the value AND the index in the match_array variable
                ' Debug.Print "We have a match: " & stringArray(j, 1)
                match_array(y, 1) = stringArray(j, 1)
                match_array(y, 2) = x
                y = y + 1
                Exit For
            End If
        Next j
    Next x
    
    ' Print match_array variable
    Dim temp As Integer
    For temp = 1 To lastRowA
        'Debug.Print "Match: " & match_array(temp, 1), match_array(temp, 2)
    Next temp
    
    ' Copy values in Column 'Masse''
    Dim match_row  As Integer, match_next_row  As Integer, nbrrows As Integer, from_to As Integer
    Dim m As Variant
    Dim match As Integer
    Dim copy_match As Integer
    
    For match = 1 To 30 ' lastRowA
        match_row = match_array(match, 2)
        match_next_row = match_array(match + 1, 2)
        from_to = match_array(match + 1, 2) - match_array(match, 2) - 1
        Debug.Print "Match row: " & match_row
        Debug.Print "Match next row: " & match_next_row
        Debug.Print "From To: " & from_to
            For copy_match = 1 To from_to
                wstA.Cells(match_row + copy_match, colIndexB + 1) = wstA.Cells(match_row + copy_match, colIndexB)
            Next copy_match
    Next match
End Sub

Public Function checkcolumnMasse(headersA As Range) As Boolean

For Each Header In headersA
        If Header = "Masse" Then
            'Debug.Print "We have a match: " & Header
            checkcolumnMasse = True
        End If
    Next Header
    
End Function
