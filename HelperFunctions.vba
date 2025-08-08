' This is in a regular module class

Option Explicit

'=== Column number to letter ===
Public Function GetColumnLetter(ByVal colNum As Long) As String
    GetColumnLetter = Replace(Cells(1, colNum).Address(False, False), "1", "")
End Function

'=== Column letter to number ===
Public Function GetColumnNum(ByVal colLetter As String) As Long
    Dim i As Long
    Dim result As Long
    Dim letter As String
    
    colLetter = UCase(colLetter) ' convert to uppercase string
    
    result = 0
    For i = 1 To Len(colLetter)
        letter = Mid(colLetter, i, 1)
        result = result * 26 + (Asc(letter) - Asc("A") + 1)
    Next i
    
    GetColumnNum = result
End Function

Public Sub ReadData
    Dim wsAlberta As Worksheet
    Dim wsScripting As Worksheet

    Dim startingRow As Integer
    Dim teamMembersQuantity As Integer
    Dim blockHeight As Integer
    Dim blockLength As Integer

    Dim projectName As String
    Dim projectLead As String
    Dim projectNumber As Variant

    Dim projectList As New Scripting.Dictionary
    Dim keyArray() As String

    Dim emptyCounter As Integer
    Dim isEndofList As Boolean

    Set wsAlberta = Worksheets("Alberta")
    Set wasScripting = Worksheets("Scripting")

    Set startingRow = wsScripting.Range("5").Value
    Set teamMembersQuantity = wsScripting.Range("B2").Value
    Set blockHeight = wsScripting.Range("B3").Value
    Set blockLength = wsScripting.Range("B4").Value

    EndofList = False

    While EndofList == False
        If Cells(i, 1).Value == "" AND Cells(i + 1, 1).Value == "" AND Cells(i + 2).Value == ""
            EndofList == True
        End If
        Else If
            For i = startingRow To 2000 Step blockHeight
            projectName = Cells(i, 1).Value
            projectLead = Cells(i + 1, 1).Value
            projectNumber = Cells(i + 3, 1).Value
            headRow = i
            ' blockheight defined
            ' blocklength defined
            ' worksheet defined

            Dim project As ProjectBlockClass
            Set project = New ProjectBlockClass
            project.Constructor projectName, projectLead, projectNumber, headRow, blockHeight, blockLength, wsAlberta
            projectList.Add project.ProjectName, project
        End If
    Next i
    
    ReDim keyArray(1 to projectList.Count) As String

    For each key in projectList
        keyArray(key) == projectList(key)
    End For
End Sub
