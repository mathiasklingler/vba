Attribute VB_Name = "T3_checkUnitType"
Sub checkUnitType()
    Dim wst2 As Worksheet
    Dim cell As Range
    Dim colNameB As String, colNameE As String, colNameM As String, colNameL As String, columnHeader As String
    Dim menge As Variant, headerRange As Variant
    Dim x As Integer, lastColumnT2 As Integer, colIndexE As Integer, colIndexM As Integer, colIndexL As Integer, colIndexB As Integer, lastRowT2 As Integer
    Dim multiplied As Long
    
    ' Set reference to the target sheet
    Set wst2 = ThisWorkbook.Worksheets("T2")
    
    ' Get lastcolumn and lastrow from wst2
    lastColumnT2 = wst2.Cells(1, wst2.Columns.Count).End(xlToLeft).Column
    lastRowT2 = wst2.Cells(wst2.Rows.Count, "D").End(xlUp).Row
    
    ' Define header range
    Set headerRange = wst2.Range("A1:Z1")
    'Set headerRange = wst2.Range("A1:A1" & headerRange)
    Debug.Print "headerRange: " & headerRange(15)
    
    ' Set header name
    colNameE = "Einheit"
    colNameM = "Menge"
    colNameL = "L"
    colNameB = "Bezeichnung_a"
       
    ' Get column index defined by header name
    colIndexE = getColIndex(headerRange, colNameE, lastColumnT2)
    colIndexM = getColIndex(headerRange, colNameM, lastColumnT2)
    colIndexL = getColIndex(headerRange, colNameL, lastColumnT2)
    colIndexB = getColIndex(headerRange, colNameB, lastColumnT2)
    Debug.Print "ColIndex Menge: " & colIndexM
    
    ' Loop through wst2 - check/edit Columns 'Menge', 'Einheit', 'L'
    For x = 2 To lastRowT2
        ' Debug.Print "col M: " & wst2.Cells(x, colIndexM)
        If wst2.Cells(x, colIndexE) = "m" And wst2.Cells(x, colIndexM) <> "" And wst2.Cells(x, colIndexM) <> 1 And wst2.Cells(x, 25) <> "Kanal" Then
        'If wst2.Cells(x, colIndexE) = "m" And wst2.Cells(x, colIndexM) <> "" And wst2.Cells(x, colIndexM) <> 1 Then
            'Debug.Print "col M: " & wst2.Cells(x, colIndexM)
            multiplied = wst2.Cells(x, colIndexM) * 1000
            wst2.Cells(x, colIndexL).Value = multiplied
            'wst2.Cells(x, colIndexM).Copy wst2.Cells(x, colIndexL)
            wst2.Cells(x, colIndexM) = 1
        End If
        If wst2.Cells(x, 25).Value = "Kanal" Then
            multiplied = wst2.Cells(x, colIndexM) * 1000
            wst2.Cells(x, colIndexL).Value = multiplied
            wst2.Cells(x, colIndexM) = 1
        End If
    Next x
        
    
End Sub


