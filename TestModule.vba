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
    Set inputCellweek = wsScripting.Range("J3")
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

Sub FindTeamMemberHoursButton()
    RefreshData
    
    Dim wsScripting As Worksheet
    Dim teamMemberName As String
    Dim inputCellteamMemberName As Range
    Dim project As String
    Dim inputCellproject As Range
    Dim week As Integer
    Dim inputCellweek As Range

    Dim outputCell As Range

    Set wsScripting = Worksheets("Scripting")
    Set inputCellteamMemberName = wsScripting.Range("J24")
    Set inputCellproject = wsScripting.Range("J25")
    Set inputCellweek = wsScripting.Range("J26")

    Set outputCell = wsScripting.Range("J27")

    If inputCellteamMemberName.value = "" Then
        outputCell.value = ""
        Exit Sub
    Else
        teamMemberName = inputCellteamMemberName.value
    End If

    If inputCellproject.value = "" Then
        outputCell.value = ""
        Exit Sub
    Else
        project = inputCellproject.value
    End If

    If inputCellweek.value = "" Then
        outputCell.value = ""
        Exit Sub
    Else
        week = inputCellweek.value
    End If

    outputCell.Value = Projects(project).GetTeamMemberHours(teamMemberName, week, team)
End Sub

Sub SetTeamMemberHoursButton()
    RefreshData
    
    Dim wsScripting As Worksheet
    Dim teamMemberName As String
    Dim inputCellteamMemberName As Range
    Dim project As String
    Dim inputCellproject As Range
    Dim week As Integer
    Dim inputCellweek As Range
    Dim hours As Integer
    Dim inputCellhours As Range
    Dim feedbackCell As Range

    Dim outputCell As Range

    Set wsScripting = Worksheets("Scripting")
    Set inputCellteamMemberName = wsScripting.Range("J34")
    Set inputCellproject = wsScripting.Range("J35")
    Set inputCellweek = wsScripting.Range("J36")
    Set inputCellhours = wsScripting.Range("J37")
    Set feedbackCell = wsScripting.Range("L37")

    Set outputCell = wsScripting.Range("J26")

    If inputCellteamMemberName.value = "" Then
        feedbackCell.Value = "Hours not set."
        Exit Sub
    Else
        teamMemberName = inputCellteamMemberName.value
    End If

    If inputCellproject.value = "" Then
        feedbackCell.Value = "Hours not set."
        Exit Sub
    Else
        project = inputCellproject.value
    End If

    If inputCellweek.value = "" Then
        feedbackCell.Value = "Hours not set."
        Exit Sub
    Else
        week = inputCellweek.value
    End If

    If inputCellhours.value = "" Then
        feedbackCell.value = "Hours not set."
        Exit Sub
    Else
        hours = inputCellhours.value
    End If

    Projects(project).SetTeamMemberHours hours, teamMemberName, week, team
    feedbackCell.Value = "Hours set."
End Sub
