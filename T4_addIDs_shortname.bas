Attribute VB_Name = "T4_addIDs_shortname"
Sub SetIDsAndShortNames()
    Dim wst2 As Worksheet, asSheet As Worksheet
    Dim headerRange As Variant
    Dim L As Long
    Dim lastColumnT2 As Integer, colIndexB As Integer, colIndexL As Integer, colIndexW As Integer
    Dim W As String, Ausführung As String, Material As String, colNameB As String, colNameL As String, colNameW As String
    
    ' Set reference to the "T2" and "Artikelstamm" worksheet
    Set wst2 = ThisWorkbook.Worksheets("T2")
    Set asSheet = ThisWorkbook.Sheets("Artikelstamm")
    
    ' Get the string array from "Artikelstamm" worksheet
    stringArray = asSheet.Range("A1:C" & asSheet.Cells(asSheet.Rows.Count, "A").End(xlUp).Row).Value
    
    ' Get lastrow and lastcolumn from wst2'
    lastRowT2 = wst2.Cells(wst2.Rows.Count, "D").End(xlUp).Row
    lastColumnT2 = wst2.Cells(1, wst2.Columns.Count).End(xlToLeft).Column
    
    ' Define header range
    Set headerRange = wst2.Range("A1:Z1")
    
    ' Set header name for column and get columnIndex
    colNameB = "Bezeichnung_a"
    colNameL = "L"
    colNameW = "W"
    colIndexB = getColIndex(headerRange, colNameB, lastColumnT2)
    colIndexL = getColIndex(headerRange, colNameL, lastColumnT2)
    colIndexW = getColIndex(headerRange, colNameW, lastColumnT2)
    
    ' Loop through each row and check if the Value in Column 'Bezeichnung'("D")
    ' in row x matches one of the strings in the array in the "Artikelstamm" table'
    
    'Loop through table
    For x = 2 To lastRowT2
        sourceString = wst2.Cells(x, colIndexB)
        'Check exceptions
        If sourceString = "Regenhut" Or sourceString = "Kanal" Or sourceString = "Bogen" Then
            L = wst2.Cells(x, colIndexL).Value
            W = wst2.Cells(x, colIndexW).Value
            Ausführung = wst2.Cells(x, 16).Value
            Material = wst2.Cells(x, 17).Value
            
            ' Set Case statement for special Articles'
            Select Case sourceString  ' Evaluate Cell Value in Column Q
                ' Check the paramater of the Regenhut'
                Case "Regenhut"
                Select Case Material
                    Case "Alu"
                    Select Case Ausführung
                        Case "eckig"
                            Debug.Print "This is an eckiger Alu Regenhut"
                            wst2.Cells(x, 1).Value = "330.3"
                            wst2.Cells(x, 2).Value = "RHVZK"
                        Case Else
                        Debug.Print "This is not an eckiger Alu Regenhut"
                        wst2.Cells(x, 1) = "bitte prüfen"
                    End Select
                    Case Else
                    Debug.Print "This is not an Alu Regenhut"
                    wst2.Cells(x, 1) = "bitte prüfen"
                End Select
                
                ' Check the paramater of the Kanal'
                Case "Kanal"
                Debug.Print "L is: " & L
                Select Case L
                    Case 3 To 222222222
                        Debug.Print "Der Kanal ist länger als 3m"
                        'wsT2.Cells(x, 1) = "9"
                        wst2.Cells(x, 1).Value = "9"
                        wst2.Cells(x, 2).Value = "NK18"
                    Case Else
                        wst2.Cells(x, 1) = "bitte prüfen"
                End Select
                
                ' Check the paramater of the Bogen'
                Case "Bogen"
                Select Case W
                    Case "90°"
                        Debug.Print "Der Bogen ist 90°"
                        wst2.Cells(x, 1) = "510"
                        wst2.Cells(x, 2) = "B90"
                    Case "60°"
                        Debug.Print "Der Bogen ist 60°"
                        wst2.Cells(x, 1) = "511"
                        wst2.Cells(x, 2) = "B60"
                    Case "45°"
                        Debug.Print "Der Bogen ist 45°"
                        wst2.Cells(x, 1) = "512"
                        wst2.Cells(x, 2) = "B45"
                    Case "30°"
                        Debug.Print "Der Bogen ist 30°"
                        wst2.Cells(x, 1) = "513"
                        wst2.Cells(x, 2) = "B30"
                    Case "15°"
                        Debug.Print "Der Bogen ist 15°"
                        wst2.Cells(x, 1) = "514"
                        wst2.Cells(x, 2) = "B15"
                    Case Else
                        wst2.Cells(x, 1) = "bitte prüfen"
                End Select
            Case Else    ' Other values.
            Debug.Print "This article is not covered in the case statements"
            wst2.Cells(x, 1) = "bitte prüfen"
            End Select
        End If
        
        'If no exception add Art-Nr. and Kurzbez from Artikelstamm worksheet
        If sourceString <> "Regenhut" And sourceString <> "Kanal" And sourceString <> "Bogen" Then
        'Debug.Print "sourceString: " & sourceString
        For j = LBound(stringArray, 1) To UBound(stringArray, 1)
            If wst2.Cells(x, colIndexB).Value = stringArray(j, 1) Then
                wst2.Cells(x, 1) = stringArray(j, 2)
                wst2.Cells(x, 2) = stringArray(j, 3)
                Exit For
            End If
        Next j
        End If
    Next x
End Sub

