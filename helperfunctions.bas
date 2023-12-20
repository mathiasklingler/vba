Attribute VB_Name = "helperfunctions"
Function StripChar(Txt As Variant) As String
With CreateObject("VBScript.RegExp")
.Global = True
.Pattern = "\D"
StripChar = .Replace(Txt, "")
End With
End Function


Function getColIndex(wsheader As Variant, colName As String, lastCol As Integer) As Integer
    ' Get header name and return columIndex
    For x = 1 To lastCol
        If wsheader.Cells(1, x) = colName Then
            getColName = wsheader.Cells(1, x).Column
        End If
    Next x
    getColIndex = getColName
End Function


