Attribute VB_Name = "T1_Combine10Tables"
Sub CombineTables()
    Dim wsFinal As Worksheet, wsA As Worksheet, wsB As Worksheet, wsC As Worksheet, wsD As Worksheet, wsE As Worksheet, wsF As Worksheet, wsG As Worksheet, wsH As Worksheet, wsI As Worksheet
    Dim lastRowFinal As Long, lastRowA As Long, lastRowB As Long, lastRowC As Long, lastRowD As Long, lastRowE As Long, lastRowF As Long, lastRowG As Long, lastRowH As Long, lastRowI As Long
    Dim dataRangeFinal As Range, dataRangeA As Range, dataRangeB As Range, dataRangeC As Range, dataRangeD As Range, dataRangeE As Range, dataRangeF As Range, dataRangeG As Range, dataRangeH As Range, dataRangeI As Range
    Dim combinedRange As Range
    Dim newSheet As Worksheet
    Dim headersFinal As Range, headersA As Range, headersB As Range, headersC As Range, headersD As Range, headersE As Range, headersF As Range, headersG As Range, headersH As Range, headersI As Range
    Dim headerCellFinal As Range, headerCellA As Range, headerCellB As Range, headerCellC As Range, headerCellD As Range, headerCellE As Range, headerCellF As Range, headerCellG As Range, headerCellH As Range, headerCellI As Range
    Dim colIndexFinal As Long, colIndexlA As Long, colIndexB As Long, colIndexC As Long, colIndexD As Long, colIndexE As Long, colIndexF As Long, colIndexG As Long, colIndexH As Long, colIndexI As Long
    Dim targetRow As Long
    Dim i As Long, j As Long
    
    ' Set references to worksheets
    Set wsFinal = ThisWorkbook.Sheets("FinaleListe")
    Set wsA = ThisWorkbook.Sheets("A")
    Set wsB = ThisWorkbook.Sheets("B")
    Set wsC = ThisWorkbook.Sheets("C")
    Set wsD = ThisWorkbook.Sheets("D")
    Set wsE = ThisWorkbook.Sheets("E")
    Set wsF = ThisWorkbook.Sheets("F")
    Set wsG = ThisWorkbook.Sheets("G")
    Set wsH = ThisWorkbook.Sheets("H")
    Set wsI = ThisWorkbook.Sheets("I")
    
    ' Find last rows in each sheet
    lastRowFinal = wsFinal.Cells(wsFinal.Rows.Count, "B").End(xlUp).Row
    lastRowA = wsA.Cells(wsA.Rows.Count, "B").End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
    lastRowC = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).Row
    lastRowD = wsD.Cells(wsD.Rows.Count, "B").End(xlUp).Row
    lastRowE = wsE.Cells(wsE.Rows.Count, "B").End(xlUp).Row
    lastRowF = wsF.Cells(wsF.Rows.Count, "B").End(xlUp).Row
    lastRowG = wsG.Cells(wsG.Rows.Count, "B").End(xlUp).Row
    lastRowH = wsH.Cells(wsH.Rows.Count, "B").End(xlUp).Row
    lastRowI = wsI.Cells(wsI.Rows.Count, "B").End(xlUp).Row
    
    ' Define ranges for data in each sheet including headers
    Set headersFinal = wsFinal.Range("A1").Resize(1, wsFinal.Cells(1, wsFinal.Columns.Count).End(xlToLeft).Column)
    Set headersA = wsA.Range("A1").Resize(1, wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column)
    Set headersB = wsB.Range("A1").Resize(1, wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column)
    Set headersC = wsC.Range("A1").Resize(1, wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column)
    Set headersD = wsD.Range("A1").Resize(1, wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column)
    Set headersE = wsE.Range("A1").Resize(1, wsE.Cells(1, wsE.Columns.Count).End(xlToLeft).Column)
    Set headersF = wsF.Range("A1").Resize(1, wsF.Cells(1, wsF.Columns.Count).End(xlToLeft).Column)
    Set headersG = wsG.Range("A1").Resize(1, wsG.Cells(1, wsG.Columns.Count).End(xlToLeft).Column)
    Set headersH = wsH.Range("A1").Resize(1, wsH.Cells(1, wsH.Columns.Count).End(xlToLeft).Column)
    Set headersI = wsI.Range("A1").Resize(1, wsI.Cells(1, wsI.Columns.Count).End(xlToLeft).Column)
    
    Set dataRangeA = wsA.Range("A2").Resize(lastRowA - 1, headersA.Columns.Count)
    Set dataRangeB = wsB.Range("A2").Resize(lastRowB - 1, headersB.Columns.Count)
    Set dataRangeC = wsC.Range("A2").Resize(lastRowC - 1, headersC.Columns.Count)
    Set dataRangeD = wsD.Range("A2").Resize(lastRowD - 1, headersD.Columns.Count)
    Set dataRangeE = wsE.Range("A2").Resize(lastRowE - 1, headersE.Columns.Count)
    Set dataRangeF = wsF.Range("A2").Resize(lastRowF - 1, headersF.Columns.Count)
    Set dataRangeG = wsG.Range("A2").Resize(lastRowG - 1, headersG.Columns.Count)
    Set dataRangeH = wsH.Range("A2").Resize(lastRowH - 1, headersH.Columns.Count)
    Set dataRangeI = wsI.Range("A2").Resize(lastRowI - 1, headersI.Columns.Count)
    
    ' Create a new sheet to display the combined data
    Set newSheet = ThisWorkbook.Worksheets("T1")
    
    ' Copy headers from sheet A to new sheet
    headersFinal.Copy Destination:=newSheet.Range("A1")
    
    ' Match columns from sheet A to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    Debug.Print "targetRowA is: " & targetRow
    For Each headerCellA In headersA
        For Each headerCellFinal In headersFinal
            If headerCellA.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexA = headerCellA.Column
                
                ' Copy data from sheet A to corresponding column in sheet Finaleliste
                wsA.Cells(1, colIndexA).Resize(lastRowA - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellA
    
    ' Match columns from sheet B to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    Debug.Print "targetRowB is: " & targetRow
    For Each headerCellB In headersB
        For Each headerCellFinal In headersFinal
            If headerCellB.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexB = headerCellB.Column
                
                ' Copy data from sheet B to corresponding column in sheet T1
                wsB.Cells(1, colIndexB).Resize(lastRowB - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellB
    
    ' Match columns from sheet C to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    Debug.Print "targetRowC is: " & targetRow
    For Each headerCellC In headersC
        For Each headerCellFinal In headersFinal
            If headerCellC.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexC = headerCellC.Column
                
                ' Copy data from sheet C to corresponding column in sheet Finaleliste
                wsC.Cells(1, colIndexC).Resize(lastRowC - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellC
    
    ' Match columns from sheet D to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    Debug.Print "targetRowC is: " & targetRow
    For Each headerCellD In headersD
        For Each headerCellFinal In headersFinal
            If headerCellD.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexD = headerCellD.Column
                
                ' Copy data from sheet D to corresponding column in sheet Finaleliste
                wsD.Cells(1, colIndexD).Resize(lastRowD - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellD
    
    ' Match columns from sheet E to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    For Each headerCellE In headersE
        For Each headerCellFinal In headersFinal
            If headerCellE.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexE = headerCellE.Column
                
                ' Copy data from sheet E to corresponding column in sheet Finaleliste
                wsE.Cells(1, colIndexE).Resize(lastRowE - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellE
    
    ' Match columns from sheet F to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    For Each headerCellF In headersF
        For Each headerCellFinal In headersFinal
            If headerCellF.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexF = headerCellF.Column
                              
                ' Copy data from sheet F to corresponding column in sheet Finaleliste
                wsF.Cells(1, colIndexF).Resize(lastRowF - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellF
    
    ' Match columns from sheet G to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    For Each headerCellG In headersG
        For Each headerCellFinal In headersFinal
            If headerCellG.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexG = headerCellG.Column
                
                ' Copy data from sheet G to corresponding column in sheet Finaleliste
                wsG.Cells(1, colIndexG).Resize(lastRowG - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellG
    
    ' Match columns from sheet H to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    For Each headerCellH In headersH
        For Each headerCellFinal In headersFinal
            If headerCellH.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexH = headerCellH.Column
                
                ' Copy data from sheet H to corresponding column in sheet Finaleliste
                wsH.Cells(1, colIndexH).Resize(lastRowH - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellH
    
    ' Match columns from sheet I to columns in sheet FinaleListe by header names and copy the data
    targetRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row + 1
    For Each headerCellI In headersI
        For Each headerCellFinal In headersFinal
            If headerCellI.Value = headerCellFinal.Value Then
                colIndexlFinal = headerCellFinal.Column
                colIndexI = headerCellI.Column
                
                ' Copy data from sheet I to corresponding column in sheet Finaleliste
                wsI.Cells(1, colIndexI).Resize(lastRowI - 1).Copy Destination:=newSheet.Cells(targetRow, colIndexlFinal)
                
                Exit For
            End If
        Next headerCellFinal
    Next headerCellI
    
    ' Autofit columns in the new sheet
    newSheet.Cells.EntireColumn.AutoFit
End Sub


