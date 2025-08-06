' This is in a regular module class

Option Explicit

'=== Column number to letter ===
Public Function GetColumnLetter(colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address(False, False), "$")(0)
End Function

'=== Column letter to number ===
Public Function GetColumnNum(colLetter As String) As Long
    Dim i As Long, result As Long
    
    colLetter = UCase(colLetter)
    
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(Mid(colLetter, i, 1)) - Asc("A") + 1)
    Next i
    
    ColNumber = result
End Function
