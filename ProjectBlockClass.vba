' === ProjectBlock.vba ===
' This is in a class module
' Strong type things like in C#
Option Explicit

'--- Properties ---
Private v_projectName As String
Private v_teamLead As String
Private v_projectNumber As Variant
Private v_projectStatus As Variant 
Private v_mainNotes As String
Private v_notes(1 To 13) As String
Private v_headRow As Integer
Private v_endRow As Integer
Private v_blockHeight As Integer
Private v_blockLength As Integer
Private v_endColLetter As String
Private v_data As Variant
Private v_ws As Worksheet

'--- Constructor-like method ---
Public Sub Constructor(ByVal a_projectName As String, _
                        ByVal a_teamLead As String, _
                        ByVal a_projectStatus As String, _
                        ByVal a_mainNotes As String, _
                        ByRef a_notes() As String, _
                        ByVal a_projectNumber as Variant, _
                        ByVal a_headRow as Integer, _
                        ByVal a_blockHeight As Integer, _
                        ByVal a_blockLength As Integer, _
                        ByVal a_worksheet As Worksheet)
    Dim i As Long
    
    v_projectName = a_projectName
    v_teamLead = a_teamLead
    v_projectStatus = a_projectStatus
    v_mainNotes = a_mainNotes

    'Setting notes array 
    If LBound(a_notes) <> 1 Or UBound(a_notes) <> 13 Then
        Err.Raise vbObjectError + 1000, , "a_notes must be an array with 13 elements"
    Else
        For i = 1 To 13
            v_notes(i) = a_notes(i)
        Next i
    End If

    v_projectNumber = a_projectNumber
    v_headRow = a_headRow
    v_blockHeight = a_blockHeight
    v_endRow = a_headRow + a_blockHeight - 1
    v_blockLength = a_blockLength
    v_endColLetter = GetColumnLetter(a_blockLength)
    Set v_ws = a_worksheet
    OccupyData
End Sub

'--- Getters / Setters ---
Public Property Get ProjectName() As String
    ProjectName = v_projectName
End Property

Public Property Let ProjectName(ByVal value As String)
    v_projectName = value
End Property

Public Property Get TeamLead() As String
    TeamLead = v_teamLead
End Property

Public Property Let TeamLead(ByVal value As String)
    v_teamLead = value
End Property

Public Property Get ProjectStatus() As String
    ProjectStatus = v_projectStatus
End Property

Public Property Let ProjectStatus(ByVal value As String)
    v_projectStatus = value
End Property

Public Property Get ProjectNumber() As Variant
    ProjectNumber = v_projectNumber
End Property

Public Property Get MainNotes() As String
    MainNotes = v_mainNotes
End Property

Public Property Let MainNotes(value As String)
    v_mainNotes = value
End Property

Public Property Get Notes(ByVal index As Integer) As String
    Notes = v_notes(index)
End Property

Public Property Let Notes(ByVal index As Integer, value As String) 
    v_notes(index) = value
End Property

Public Property Let ProjectNumber(ByVal value as Variant)
    v_projectNumber = value
End Property

Public Property Get HeadRow() As Integer
    HeadRow = v_headRow
End Property

Public Property Let HeadRow(ByVal value As Integer)
    v_headRow = value
    v_endRow = v_headRow + v_blockHeight - 1
End Property

Public Property Get BlockHeight() As Integer
    BlockHeight = v_blockHeight
End Property

Public Property Let BlockHeight(ByVal value As Integer)
    v_blockHeight = value
    v_endRow = v_headRow + v_blockHeight - 1
End Property

' Add project block method to insert the project block onto a worksheet
Public Sub AddProjectBlock(team As TeamMembers, templateSheet As String)

    'Comment out. For debugging
    Dim v_ws As Worksheet
    Set v_ws = Worksheets("Test")

    ' Insert Rows
    Dim i as Integer
    For i = 1 to v_blockHeight 
        v_ws.Rows(v_headRow).Insert Shift:=xlDown
        Debug.Print v_headRow
    Next i

    InsertData

    ' Populate all items on the left block: Project Name, tead lead, project number, main notes, notes
    v_ws.Cells(v_headRow, "A").Value = v_projectName
    v_ws.Cells(v_headRow + 1, 1).Value = v_teamLead
    v_ws.Cells(v_headRow + 2, 1).Value = v_projectNumber
    v_ws.Cells(v_headRow + 3, 1).Value = v_mainNotes
    v_ws.Cells(V_headRow + 4, 1).Value = "Project Status:"
    For i = 1 To 13
        v_ws.Cells(v_headRow + 7 + i, 1).Value = v_notes(i)
    Next i
    
    ' Create data validation for the cells for project status. The data validation range is hardcoded from 1:100 on a worksheet called "List"
    With v_ws.Cells(v_headRow + 5, 1).Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlEqual, _
            Formula1:="=List!A1:A100"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' Add project status
    v_ws.Cells(v_headRow + 5, 1).Value = v_projectStatus

    ' Populate Team 
    v_ws.Cells(v_headRow, 2).Value = "*"
    For i = 1 to v_blockHeight - 1
        v_ws.Cells(v_headRow + i, 2).Value = team.TeamMembersNum(i)
    Next i

    ' Add sum formulas on the last column of the project block
    For i = 1 to v_blockHeight
        v_ws.Cells(v_headRow + i, v_blockLength).Formula = _
            "=SUM(C" & (v_headRow + i) & ":" & GetColumnLetter(v_blockLength - 1) & (v_headRow + i) & ")"
    Next i

    ' Formatting Copy/Paste from Template
    Dim templateRange As Range
    Dim desRange As Range
    Set templateRange = Worksheets(templateSheet).Range("A1:" & v_endColLetter & v_blockLength)
    Set desRange = v_ws.Range("A" & v_headRow & ":" & v_endColLetter & v_endRow)
    templateRange.Copy
    desRange.PasteSpecial xlPasteFormats

    ' Set Widths
    v_ws.columns(1).ColumnWidth = 64.14
    v_ws.columns(2).ColumnWidth = 11
    for i = 3 to v_blockLength
        v_ws.columns(i).ColumnWidth = 10
    Next i
    for i = v_headRow to v_endRow
        v_ws.rows(i).RowHeight = 15
    Next i
End Sub

Public Sub DeleteProject()
    Dim deleteRange as Range
    Set deleteRange = v_ws.Range("A" & v_headRow & ":A" & v_endRow)
    deleteRange.EntireRow.Delete
End Sub

Public Sub OccupyData()
    ' Occupy data. v_headRow + 1 is to match TeamMembers index starting at the 2nd row of the project block. v_blocklength - 1 is to remove
    ' the column at the end which is a summation 
    v_data = v_ws.Range("C" & (v_headRow + 1) & ":" & GetColumnLetter(v_blockLength - 1) & v_endRow)
End Sub

Public Sub InsertData()
    Dim i As Long, j As Long

    ' For Debugging
    Set v_ws = Worksheets("Test")

    ' Outer if statement checks if v_data 2D array has been made. If so then writes it to cell. If not then writes blanks
    If IsArray(v_data) And Not IsEmpty(v_data) Then
        For i = LBound(v_data, 1) to UBound(v_data, 1)
            For j = LBound(v_data, 2) to UBound(v_data, 2)
                If v_data(i,j) = 0 Then
                    v_ws.Cells(v_headrow + i, GetColumnNum("C") + j - 1).Value = ""
                Else
                    v_ws.Cells(v_headrow + i, GetColumnNum("C") + j - 1).Value = v_data(i,j)
                End If
            Next j
        Next i
    Else
        For i = LBound(v_data, 1) to UBound(v_data, 1)
            For j = LBound(v_data, 2) to UBound(v_data, 2)
                v_ws.Cells(v_headrow + i, GetColumnNum("C") + j - 1).Value = ""
            Next j
        Next i
        OccupyData
    End If
End Sub

Public Function GetTeamMemberHours(ByVal teamMemberName As String, ByVal week as Integer, ByVal team as TeamMembers) As Integer
    If Not IsEmpty(v_data) Then
        GetTeamMemberHours = v_data(team.TeamMembersName(teamMemberName), week)
    End If
End Function

Public Sub SetTeamMemberHours(ByVal hours as Variant, ByVal teamMemberName As String, ByVal week as Integer, ByVal team as TeamMembers)
    If Not IsEmpty(v_data) Then
        v_data(team.TeamMembersName(teamMemberName), week) = hours
        v_ws.Cells(headRow + team.TeamMembersName(teamMemberName), 2 + week).Value = hours
    End If
End Sub

' Assume v_notes(1 To 13) is declared at the class level

' === Getter Methods ===
Public Function GetNotes1() As String: GetNotes1 = v_notes(1): End Function
Public Function GetNotes2() As String: GetNotes2 = v_notes(2): End Function
Public Function GetNotes3() As String: GetNotes3 = v_notes(3): End Function
Public Function GetNotes4() As String: GetNotes4 = v_notes(4): End Function
Public Function GetNotes5() As String: GetNotes5 = v_notes(5): End Function
Public Function GetNotes6() As String: GetNotes6 = v_notes(6): End Function
Public Function GetNotes7() As String: GetNotes7 = v_notes(7): End Function
Public Function GetNotes8() As String: GetNotes8 = v_notes(8): End Function
Public Function GetNotes9() As String: GetNotes9 = v_notes(9): End Function
Public Function GetNotes10() As String: GetNotes10 = v_notes(10): End Function
Public Function GetNotes11() As String: GetNotes11 = v_notes(11): End Function
Public Function GetNotes12() As String: GetNotes12 = v_notes(12): End Function
Public Function GetNotes13() As String: GetNotes13 = v_notes(13): End Function

' === Setter Methods ===
Public Sub SetNotes1(value As String): v_notes(1) = value: End Sub
Public Sub SetNotes2(value As String): v_notes(2) = value: End Sub
Public Sub SetNotes3(value As String): v_notes(3) = value: End Sub
Public Sub SetNotes4(value As String): v_notes(4) = value: End Sub
Public Sub SetNotes5(value As String): v_notes(5) = value: End Sub
Public Sub SetNotes6(value As String): v_notes(6) = value: End Sub
Public Sub SetNotes7(value As String): v_notes(7) = value: End Sub
Public Sub SetNotes8(value As String): v_notes(8) = value: End Sub
Public Sub SetNotes9(value As String): v_notes(9) = value: End Sub
Public Sub SetNotes10(value As String): v_notes(10) = value: End Sub
Public Sub SetNotes11(value As String): v_notes(11) = value: End Sub
Public Sub SetNotes12(value As String): v_notes(12) = value: End Sub
Public Sub SetNotes13(value As String): v_notes(13) = value: End Sub