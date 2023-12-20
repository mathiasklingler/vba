Attribute VB_Name = "Transformation2_CombineTables"
Sub CombineTables()
    Dim wsA As Worksheet, wsB As Worksheet
    Dim lastRowA As Long, lastRowB As Long
    Dim dataRangeA As Range, dataRangeB As Range
    Dim combinedRange As Range
    Dim newSheet As Worksheet
    Dim headersA As Range, headersB As Range
    Dim headerCellA As Range, headerCellB As Range
    Dim colIndexA As Long, colIndexB As Long
    Dim targetRow As Long
    Dim i As Long, j As Long
    
    ' Set references to worksheets
    Set wsA = ThisWorkbook.Sheets("FinaleListe")
    Set wsB = ThisWorkbook.Sheets("B")
    Set wsC = ThisWorkbook.Sheets("C")
    
    ' Find last rows in each sheet
    lastRowA = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
    lastRowC = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).Row
    
    ' Define ranges for data in each sheet including headers
    Set headersA = wsA.Range("A1").Resize(1, wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column)
    Set headersB = wsB.Range("A1").Resize(1, wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column)
    Set headersC = wsC.Range("A1").Resize(1, wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column)
    
    Set dataRangeB = wsB.Range("A2").Resize(lastRowB - 1, headersB.Columns.Count)
    Set dataRangeC = wsC.Range("A2").Resize(lastRowC - 1, headersC.Columns.Count)
    
    ' Create a new sheet to display the combined data
    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    ' Copy headers from sheet A to new sheet
    headersA.Copy Destination:=newSheet.Range("A1")
    
    ' Match columns from sheet B to columns in sheet A by header names and copy the data
    For Each headerCellB In headersB
        For Each headerCellA In headersA
            If headerCellB.Value = headerCellA.Value Then
                colIndexA = headerCellA.Column
                colIndexB = headerCellB.Column
                
                ' Find the next available row in the target column
                targetRow = newSheet.Cells(newSheet.Rows.Count, colIndexA).End(xlUp).Row + 1
                
                ' Copy data from sheet B to corresponding column in sheet A
                wsB.Cells(1, colIndexB).Resize(lastRowB - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexA)
                
                Exit For
            End If
        Next headerCellA
    Next headerCellB
    
    ' Match columns from sheet C to columns in sheet A by header names and copy the data
    For Each headerCellC In headersC
        For Each headerCellA In headersA
            If headerCellC.Value = headerCellA.Value Then
                colIndexA = headerCellA.Column
                colIndexC = headerCellC.Column
                
                ' Find the next available row in the target column
                targetRow = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row + 1
                'lastRowB = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row'
    
                ' Copy data from sheet C to corresponding column in sheet A
                wsC.Cells(1, colIndexC).Resize(lastRowC - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexA)
                
                Exit For
            End If
        Next headerCellA
    Next headerCellC
    
    ' Autofit columns in the new sheet
    newSheet.Cells.EntireColumn.AutoFit
End Sub


