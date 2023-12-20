Attribute VB_Name = "T2_DeleteRows"
Public Sub DeleteRows()
    Dim metadata As Worksheet, wsT1 As Worksheet, wst2 As Worksheet
    Dim sourceString As String
    Dim asSheet As Worksheet
    Dim stringArray As Variant
    Dim i As Long, j As Long, lastRowT1 As Long, lastColumnT1 As Long, lastRowT2 As Long
    Dim x As Long, y As Long
    Dim cell As String
    Dim match As String
    
    ' Set reference to the "T1" and "Artikelstamm" worksheet
    Set wsT1 = ThisWorkbook.Worksheets("T1")
    Set wst2 = ThisWorkbook.Worksheets("T2")
    Set asSheet = ThisWorkbook.Sheets("Artikelstamm")
    
    ' Get the string array from the specified range in "Artikelstamm" worksheet
    stringArray = asSheet.Range("A1:A" & asSheet.Cells(asSheet.Rows.Count, "A").End(xlUp).Row).Value
    
    ' Get lastrow and lastcolumn from "T1"'
    lastRowT1 = wsT1.Cells(wsT1.Rows.Count, "D").End(xlUp).Row
    lastColumnT1 = wsT1.Cells(1, wsT1.Columns.Count).End(xlToLeft).Column
    
    ' Loop through each row and check if the Value in Column 'Bezeichnung'("D")
    ' in row x matches one of the strings in the array in the "Artikelstamm" table'
    For x = 1 To lastRowT1
        sourceString = wsT1.Cells(x, 4)
        lastRowT2 = wst2.Cells(wst2.Rows.Count, "D").End(xlUp).Row + 1 ' Set to paste below last row
        For j = LBound(stringArray, 1) To UBound(stringArray, 1)
            If InStr(1, sourceString, stringArray(j, 1)) > 0 Then
                match = stringArray(j, 1)
                wsT1.Range("A" & x).EntireRow.Copy wst2.Range("A" & lastRowT2)
                wst2.Cells(lastRowT2, lastColumnT1 + 1).Value = match
                'wsT2.Cells(lastRowT2, 1) = stringArray(j, 2)'
                'wsT2.Cells(lastRowT2, 2) = stringArray(j, 3)'
                Exit For
            End If
        Next j
    Next x
End Sub
