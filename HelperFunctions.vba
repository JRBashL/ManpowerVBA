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

Public Sub ReadProjectData(ByRef projectList As Scripting.Dictionary, ByRef keyArray() As String, _
                            Optional ByVal skipProjectsArgument As Variant)
    ' Reset projectlist and keyArray for refreshing the data in runtime
    projectList.RemoveAll
    Erase keyArray
    
    ' Shorthand worksheets
    Dim wsAlberta As Worksheet
    Dim wsScripting As Worksheet

    ' Variables for skipProjects string array
    Dim projectsToSkip() As Variant

    ' Variables for the project block class
    Dim startingRow As Integer
    Dim teamMembersQuantity As Integer
    Dim blockHeight As Integer
    Dim blockLength As Integer
    Dim headRow As Integer
    Dim projectName As String
    Dim projectLead As String
    Dim projectStatus As String
    Dim mainNotes As String
    Dim notes(1 To 13) As String
    Dim projectNumber As Variant
    Dim project As ProjectBlockClass

    ' Variables for loops
    Dim emptyCounter As Integer
    Dim classListCounter As Integer
    Dim isEndOfList As Boolean
    Dim arrayCounter As Integer
    Dim key As Variant
    Dim entry As Variant
    Dim i As Variant

    Set wsAlberta = Worksheets("Alberta")
    Set wsScripting = Worksheets("Scripting")

    ' Read important variables off the scripting worksheet
    startingRow = wsScripting.Range("B5").Value
    teamMembersQuantity = wsScripting.Range("B2").Value
    blockHeight = wsScripting.Range("B3").Value
    blockLength = wsScripting.Range("B4").Value

    isEndOfList = False
    classListCounter = startingRow

    ' Check if the optional array was provided. If not, uses default. If yes then assigns to local variable
    If IsMissing(skipProjectsArgument) Then
        projectsToSkip = Array("Weekly Manpower", "% Billable", "Billable Hours","")
    ElseIf VarType(skipProjectsArgument) = vbArray + vbString Then
        projectsToSkip = skipProjectsArgument
    Else
        Err.Raise vbObjectError + 1000, "ReadProjectData", "ReadProjectData requires skipProjects argument to be a string array."
    End If
        
    ' Main while loop to read the project blocks in the "Alberta" worksheet. Creates a dictionary of projectblocks, using the project name as the key
    ' While Loop stops when it finds 3 consecutive cells to be empty in the next iteration. While loop steps by classListCounter, which is set to the projectblock height
    Do While isEndOfList = False

        projectName = wsAlberta.Cells(classListCounter, 1).Value

        If wsAlberta.Cells(classListCounter, 1).Value = "" And wsAlberta.Cells(classListCounter + 1, 1).Value = ""And wsAlberta.Cells(classListCounter + 2).Value = "" Then
            isEndOfList = True
        ElseIf CheckMatchStringArray(projectName, projectsToSkip) = True Then
            ' Skips to the next project block if the project name matches any of the elements in the projectsToSkip array.
            classListCounter = classListCounter + blockHeight
        Else
            projectLead = wsAlberta.Cells(classListCounter + 1, 1).Value
            projectStatus = wsAlberta.Cells(classListCounter + 5, 1).Value
            mainNotes = wsAlberta.Cells(classListCounter + 3, 1).Value
            projectNumber = wsAlberta.Cells(classListCounter + 2, 1).Value
            For i = 1 To 13
                notes(i) = wsAlberta.Cells(classListCounter + 7 + i, 1).Value
            Next i
            headRow = classListCounter
            ' blockheight already defined
            ' blocklength defined
            ' worksheet defined

            ' Create instance and add to list
            Set project = New ProjectBlockClass
            
            If projectList.Exists(projectName) Then
                projectName = projectName + " DUPLICATE PROJECT"              
            End if 
            project.Constructor projectName, projectLead, projectStatus, mainNotes, notes, projectNumber, headRow, blockHeight, blockLength, wsAlberta
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
        Debug.Print "Key: " & key & " | Value: " & projectList(key).ProjectName  
    Next key

    ' Print all keys stored in keyArray in Immediate window, and also writes to cells in the Worksheets("Scripting")
    wsScripting.Range("G2:G500").Value = ""
    For i = LBound(keyArray) To UBound(keyArray)
        Debug.Print "keyArray(" & i & "): " & keyArray(i)
        wsScripting.Cells(i + 1, 7).Value = keyArray(i)
    Next i
End Sub

' Function generates a string that shows how many hours a team member is working on a project at a given week. Optionally recursive calls for
' following weeks using maxWeek argument
Public Function CreateWeekReport(ByVal teamMemberName As String, _
                            ByVal week As Integer, _
                            projectList As Scripting.Dictionary, _
                            team As TeamMembers, _
                            Optional maxWeek As Integer = 0, _
                            Optional isInitialCall As Boolean = True, _
                            Optional recursionCounter As Integer = 0) As String

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
        ' This line tells VBA to ignore any errors and jump to the next line.
        On Error Resume Next
        currentHours = projectList(key).GetTeamMemberHours(teamMemberName, week, team)
        
        ' Check if an error occurred. Err.Number will be 0 if no error.
        If Err.Number <> 0 Then
            ' An error occurred. Clear it so it doesn't affect subsequent code.
            Err.Clear
            ' Use GoTo to jump to the next iteration of the loop.
            GoTo NextKey
        End If
        
        ' Reset error handling to its default state.
        On Error GoTo 0
        ' If hours are more than zero, extracts name and hours into separate collections, and tallies up total hours.
        If currentHours > 0 Then
            textList.Add projectList(key).ProjectName
            hoursList.Add currentHours
            totalhours = totalHours + currentHours
        End If
NextKey:
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
    if maxWeek > 0 And recursionCounter < maxWeek Then
        output = output + "Your hours for the following week" & vbNewLine & vbNewLine   
        output = output + CreateWeekReport(teamMemberName, week + 1, projectList, team, maxWeek, False, recursionCounter + 1)
    End If

    CreateWeekReport = output
End Function

' Function compares an input string with all of the elements in the checkArray string array. Returns false if none of the elements match the input string
Public Function CheckMatchStringArray(ByVal inputString As String, ByRef checkArray() As Variant) As Boolean
    Dim arrElement As Variant

    ' Loops through each index and compares. If any of the entries in the array is matching, exits function with true
    For Each arrElement in checkArray
        If inputString = arrElement Then
            CheckMatchStringArray = True
            Exit Function
        End If
    Next arrElement

    ' If Function fully completes, i.e. none of the elements in the checkArray matches the input, exits function with false
    CheckMatchStringArray = False
End Function

Public Sub UnlockScriptingSheet()
    Dim wsScripting As Worksheet
    Dim myPassword As String

    Set wsScripting = WorkSheets("Scripting")
    myPassword = "ManpowerVBA"
    
    ' Unlock the worksheet with the password
    If wsScripting.ProtectContents = True Then
        wsScripting.Unprotect Password:=myPassword
    End If
End Sub

Public Sub LockScriptingSheet()
    Dim wsScripting As Worksheet
    Dim myPassword As String

    Set wsScripting = WorkSheets("Scripting")
    myPassword = "ManpowerVBA"
    
    ' Unlock the worksheet with the password
    If wsScripting.ProtectContents = False Then
        wsScripting.Protect Password:=myPassword
    End If
End Sub

' Sub sorts the project list. First reads project data, removes all projects, sorts the projectskeyarray by alphabetrical order and then adds all projects
Public Sub SortProjects(ByVal a_team As TeamMembers, _
                        ByRef a_projectList As Scripting.Dictionary, _
                        ByRef a_projectKeyArray() As String, _
                        ByVal a_blockHeight As Integer, _
                        ByVal a_startingRow As Integer)
    RefreshData
    UnlockScriptingSheet

    'Declare worksheets
    Dim wsTemplate As Worksheet

    'Declare counters
    Dim i As Variant, j As Variant, k As Variant
    Dim project As Variant
    Dim counter As Long

    'Declare front mid, and rear arrays. This is for removing and placing some projects at the front of the list or at the back of the list
    Dim frontArray As Variant, midArray As Variant, rearArray As Variant

    Set wsTemplate = Worksheets("Template")

    k = a_startingRow

    ' Define front arrays and rear arrays. In the future these should be passed in by the button wrapper function. 
    ' The array values should be read off "Scripting" worksheet instead of hardcoded
    frontArray = Array("AWW", "G&A -Admin", "Vacation", "Stat", "Training", "Business Development and Marketing")
    rearArray = Array("BLANK TEMPLATE (Mechanical Lead)", "PENDING PROJECTS - YEG & YYC", "PENDING PROJECTS - YYZ")

    ' Delete projects from main worksheet
    'For each project in a_projectKeyArray
   '     project.DeleteProject
    'Loop

    ' Quicksort function on a_projecKeyArray
    QuickSortAlphabetical a_projectKeyArray, LBound(a_projectKeyArray), UBound(a_projectKeyArray)

    ' Remove the elements from the main array and place in new array
    midArray = RemoveKeysFromArray(a_projectKeyArray, frontArray, rearArray)

    ' Combine and redefine a_projectKeyArray
    ' Redimension the array to the new total size (Front + Middle + Back)
    Dim newSize As Long
    newSize = UBound(frontArray) - LBound(frontArray) + 1 + _
              UBound(midArray) - LBound(midArray) + 1 + _
              UBound(rearArray) - LBound(rearArray) + 1

    ReDim a_projectKeyArray(1 To newSize) ' Using 1-based index for simplicity

    counter = 1

    ' Populate the final array: FRONT
    For Each project In frontArray
        a_projectKeyArray(counter) = project
        counter = counter + 1
    Next project

    ' Populate the final array: MIDDLE (Alphabetically Sorted)
    For Each project In midArray
        a_projectKeyArray(counter) = project
        counter = counter + 1
    Next project

    ' Populate the final array: BACK
    For Each project In rearArray
        a_projectKeyArray(counter) = project
        counter = counter + 1
    Next project

    ' Reapply headrows according to the new sorting and adds project to the list
    For each project in a_projectKeyArray
        a_projectList(project).HeadRow = k
        a_projectList(project).AddProjectBlock team, wsTemplate
        k = k + a_blockHeight
        Debug.Print project
    Next project

    LockScriptingSheet
End Sub

' Quicksort algorithm with Hoare partition 
Public Sub QuickSortAlphabetical(ByRef a_stringArray() As String, ByVal i As Long, ByVal j As Long)
    Dim first As Long, last As Long
    Dim pivot As String
    Dim temp As String

    first = i
    last = j

    ' Create the pivot in the middle of the array. The \ is an integer division and truncates the decimal
    pivot = a_stringArray((i + j) \ 2)

    ' Partition Loop. Outer loop continues until the indeces end up in the middle (i = j) and the < is just to make sure 
    Do While i <= j
        ' Sub loop starts from the lower bound of the string and compares each to the pivot. The loop stops if it finds one that is greater than
        ' The pivot. Greater than the pivot means it's alphabatically later
        Do While StrComp(a_stringArray(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        ' Sub loop starts from the upper bound of the string and compares each to the pivot. The loop stops if it finds one that is less than
        ' The pivot. Less than the pivot means it's alphabatically earlier
        Do While StrComp(a_stringArray(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        ' If statement to run if the indeces are not in the middle. If so, then the array element marked for swapping (where the sub loops stopped)
        ' are swapped with each other. This uses one temp variable to store the value in while swapping
        If i <= j Then
            temp = a_stringArray(i)
            a_stringArray(i) = a_stringArray(j)
            a_stringArray(j) = temp
            ' Shift the indeces to avoid infinite loops in the next outer loop
            i = i + 1
            j = j - 1
        End If
    Loop

    ' Recursion for the left and right sub-arrays
    If first < j Then
        QuickSortAlphabetical a_stringArray, first, j
    End If
    If i < last Then
        QuickSortAlphabetical a_stringArray, i, last
    End If
End Sub

' Function to remove some items from an array according to two string array inputs frontKeys and backKeys
Public Function RemoveKeysFromArray(ByRef sourceArray() As String, ByRef frontKeys As Variant, ByRef backKeys As Variant) As String()
    
    Dim tempDict As New Scripting.Dictionary
    Dim item As Variant
    Dim subItem As Variant
    Dim combinedKeys As Variant
    Dim outputArray() As String
    Dim counter As Long
    
    ' 1. Load all elements of the sourceArray into a temporary Dictionary
    ' This is the easiest way to manage keys for removal
    For Each item In sourceArray
        tempDict.Add item, 1 ' Add the project name as the key
    Next item

    ' 2. Remove frontKeys and backKeys from the temporary Dictionary
    combinedKeys = Array(frontKeys, backKeys)
    
    For Each item In combinedKeys
        If IsArray(item) Then
            For Each subItem In item
                ' Check if the key exists before removing (prevents runtime error)
                If tempDict.Exists(subItem) Then
                    tempDict.Remove subItem
                End If
            Next subItem
        End If
    Next item

    ' 3. Convert the remaining Dictionary keys (the middle, sorted array) back to a string array
    ReDim outputArray(1 To tempDict.Count)
    counter = 1
    
    For Each item In tempDict.Keys
        outputArray(counter) = item
        counter = counter + 1
    Next item
    
    ' Return the array containing only the elements that will form the middle section
    RemoveKeysFromArray = outputArray
End Function