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

Public Sub ReadProjectData
    Dim wsAlberta As Worksheet
    Dim wsScripting As Worksheet

    Dim startingRow As Integer
    Dim teamMembersQuantity As Integer
    Dim blockHeight As Integer
    Dim blockLength As Integer
    Dim headRow As Integer

    Dim projectName As String
    Dim projectLead As String
    Dim projectNumber As Variant

    Dim project As ProjectBlockClass

    Dim projectList As New Scripting.Dictionary
    Dim keyArray() As String

    Dim emptyCounter As Integer
    Dim classListCounter As Integer
    Dim isEndOfList As Boolean
    Dim arrayCounter As Integer
    Dim key As Variant

    Set wsAlberta = Worksheets("Alberta")
    Set wsScripting = Worksheets("Scripting")

    startingRow = wsScripting.Range("B5").Value
    teamMembersQuantity = wsScripting.Range("B2").Value
    blockHeight = wsScripting.Range("B3").Value
    blockLength = wsScripting.Range("B4").Value

    isEndOfList = False
    classListCounter = startingRow

    Do While isEndOfList = False
        If wsAlberta.Cells(classListCounter, 1).Value = "" And wsAlberta.Cells(classListCounter + 1, 1).Value = ""And wsAlberta.Cells(classListCounter + 2).Value = "" Then
            isEndOfList = True
        Else
            projectName = wsAlberta.Cells(classListCounter, 1).Value
            projectLead = wsAlberta.Cells(classListCounter + 1, 1).Value
            projectNumber = wsAlberta.Cells(classListCounter + 3, 1).Value
            headRow = classListCounter
            ' blockheight defined
            ' blocklength defined
            ' worksheet defined

            Set project = New ProjectBlockClass
            project.Constructor projectName, projectLead, projectNumber, headRow, blockHeight, blockLength, wsAlberta

            projectList.Add project.ProjectName, project

            classListCounter = classListCounter + blockHeight
        End If
    Loop
    
    ReDim keyArray(1 to projectList.Count) As String

    arrayCounter = 1
    For each key in projectList.Keys
        keyArray(arrayCounter) = key
        arrayCounter = arrayCounter + 1
    Next key

    ' Print all dictionary entries (keys and values)
    For Each key In projectList.Keys
        Debug.Print "Key: " & key & " | Value: " & projectList(key).ProjectName  ' or any property you want
    Next key

    ' Print all keys stored in keyArray
    Dim i As Long
    For i = LBound(keyArray) To UBound(keyArray)
        Debug.Print "keyArray(" & i & "): " & keyArray(i)
    Next i
End Sub
