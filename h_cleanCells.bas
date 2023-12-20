Attribute VB_Name = "h_cleanCells"
Public Sub Cleancells()
Attribute Cleancells.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim wsT3 As Worksheet
    Dim selcetedArray As Variant
    Dim cleanedselectedArray As Variant
    Dim i As Integer, j As Long
    Dim x As Long, y As Long
    
    ' Set reference to the "T2" and "Artikelstamm" worksheet
    Set wsT3 = ThisWorkbook.Worksheets("T2")
    
    Dim myRange As Range
    Set myRange = Selection
    Debug.Print myRange.Address

    Debug.Print "stringArray.Count: " & myRange.Count
    For Each rang In myRange
        Debug.Print "stringArray is: " & rang(i + 1)
        Debug.Print "current cell: " & rang.Address
        cleanedselectedArray = StripChar(rang(i + 1))
        rang(i + 1) = cleanedselectedArray
    Next rang
End Sub


