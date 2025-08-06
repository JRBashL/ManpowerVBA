' This is in a regular module class

Option Explicit

'=== Column number to letter ===
Public Function GetColumnLetter(ByVal colNum As Long) As String
    GetColumnLetter = Split(Cells(1, colNum).Address(False, False), "$")(0)
End Function

'=== Column letter to number ===
Public Function GetColumnNum(ByVal colLetter As String) As Long
    Dim i As Long, result As Long
    
    GetColumnNum = UCase(colLetter)
    
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(Mid(colLetter, i, 1)) - Asc("A") + 1)
    Next i
    
    ColNumber = result
End Function
