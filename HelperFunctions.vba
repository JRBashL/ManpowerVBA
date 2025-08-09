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

Public Sub ReadProjectData(ByRef projectList As Scripting.Dictionary, ByRef keyArray() As String)
    ' Shorthand worksheets
    Dim wsAlberta As Worksheet
    Dim wsScripting As Worksheet

    ' Variables for the project block class
    Dim startingRow As Integer
    Dim teamMembersQuantity As Integer
    Dim blockHeight As Integer
    Dim blockLength As Integer
    Dim headRow As Integer
    Dim projectName As String
    Dim projectLead As String
    Dim projectNumber As Variant
    Dim project As ProjectBlockClass

    ' Variables for loops
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

    ' Main while loop to read the project blocks in the "Alberta" worksheet. Creates a dictionary of projectblocks, using the project name as the key
    ' While Loop stops when it finds 3 consecutive cells to be empty in the next iteration. While loop steps by classListCounter, which is set to the projectblock height
    Do While isEndOfList = False
        If wsAlberta.Cells(classListCounter, 1).Value = "" And wsAlberta.Cells(classListCounter + 1, 1).Value = ""And wsAlberta.Cells(classListCounter + 2).Value = "" Then
            isEndOfList = True
        Else
            projectName = wsAlberta.Cells(classListCounter, 1).Value
            projectLead = wsAlberta.Cells(classListCounter + 1, 1).Value
            projectNumber = wsAlberta.Cells(classListCounter + 3, 1).Value
            headRow = classListCounter
            ' blockheight already defined
            ' blocklength defined
            ' worksheet defined

            ' Create instance and add to list
            Set project = New ProjectBlockClass
            project.Constructor projectName, projectLead, projectNumber, headRow, blockHeight, blockLength, wsAlberta
            projectList.Add project.ProjectName, project

            'Next iteration to jump to the next headrow of the next projectblock
            classListCounter = classListCounter + blockHeight
        End If
    Loop

    ' Create array of project names
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

' Function generates a string that shows how many hours a team member is working on a project at a given week. Optionally recursive calls for
' following weeks using maxWeek argument
Public Function CreateWeekReport(ByVal teamMemberName As String, _
                            ByVal week As Integer, _
                            projectList As Scripting.Dictionary, _
                            team As TeamMembers, _
                            Optional maxWeek As Integer = 0, _
                            Optional isInitialCall As Boolean = True) As String

    ' Declaration for parsing the projectList and extracting relevant data
    dim key As Variant
    dim currentHours As Integer
    dim textList As Collection
    dim hoursList As Collection
    dim totalHours As Long

    ' Declaration for output string
    dim i As Integer
    dim output As String

    Set textList = New Collection
    Set hoursList = New Collection
    totalhours = 0

    ' Goes through projectList dictionary of ProjectBlockClass Instances
    For each key in projectList.Keys
        currentHours = projectList(key).GetTeamMemberHours(teamMemberName, week, team)
        ' If hours are more than zero, extracts name and hours into separate collections, and tallies up total hours.
        If currentHours > 0 Then
            textList.Add projectList(key).ProjectName
            hoursList.Add currentHours
            totalhours = totalHours + currentHours
        End If
    Next key

    ' String output generation
    If isInitialCall Then
        output = "Hi " & teamMemberName & ". Your hours for this week:" & vbNewLine & vbNewLine
    End If
    ' Appends to string for each item in the collections
    For i = 1 to textList.Count
        output = output + textList(i) & ": " & hoursList(i) & " hours." & vbNewLine
    Next i
    'Appends to string the total
    output = output + vbNewLine & "Total: " & totalHours & vbNewLine & vbNewLine

    ' Recursion for following weeks
    if maxWeek > 0 And week < maxWeek Then
        output = output + "Your hours for the following week" & vbNewLine & vbNewLine   
        output = output + CreateWeekReport(teamMemberName, week + 1, projectList, team, maxWeek, False)
    End If

    CreateWeekReport = output
End Function

