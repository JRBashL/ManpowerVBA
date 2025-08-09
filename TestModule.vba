' Test
Public team As TeamMembers
Public TestProject As ProjectBlockClass

Public Projects As New Scripting.Dictionary
Public ProjectKeyArray() As String

Sub RefreshData()
    Set team = New TeamMembers
    team.Constructor
    ReadProjectData Projects, ProjectKeyArray
End Sub

Sub WeekReportButton()

    RefreshData
    
    Dim wsScripting As Worksheet
    Dim teamMemberName As String
    Dim inputCellteamMemberName As Range
    Dim week As Integer
    Dim inputCellweek As Range
    Dim maxWeek As Integer
    Dim inputCellmaxWeek As Range
    Dim outputCell As Range

    Set wsScripting = Worksheets("Scripting")
    Set inputCellteamMemberName = wsScripting.Range("J2")
    Set inputCellweek = wsScripting.Range.("J3")
    Set inputCellmaxWeek = wsScripting.Range("J5")
    Set outputCell = wsScripting.Range("J7")
    
    
    If inputCellteamMemberName.value = "" Then
        outputCell.value = ""
        Exit Sub
    Else
        teamMemberName = inputCellteamMemberName.value
    End If
    
    If inputCellweek.value = "" Then
        outputCell.value = ""
        Exit Sub
    Else
        week = inputCellweek.value
    End If
    
    If inputCellmaxWeek.value = "" Then
        maxWeek = 0
    End If
        maxWeek = wsScripting.Range("J5").value
        
    outputCell.value = CreateWeekReport(teamMemberName, week, Projects, team, maxWeek)
End Sub
