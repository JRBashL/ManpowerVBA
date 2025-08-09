' Test
Public team As TeamMembers
Public TestProject As ProjectBlockClass

Public Projects As New Scripting.Dictionary
Public ProjectKeyArray() As String

Sub CreateTest()
    Set team = New TeamMembers
    team.Constructor

    Dim projectName As String
    Dim projectLead As String
    Dim projectNum As Long
    Dim startRow As Integer
    Dim blockheight As Integer
    Dim blockwidth As Integer
    Dim worksheetName As String

    projectName = "My Cool Project"
    projectLead = "Pertti"
    projectNum = 123456
    startRow = 55
    blockheight = team.TeamMembersNum.Count + 1
    blockwidth = 35
    worksheetName = "Test"

    Set TestProject = New ProjectBlockClass
    
    TestProject.Constructor projectName, projectLead, projectNum, startRow, blockheight, blockwidth, worksheetName
    TestProject.AddProjectBlock team, "Template"

    ReadProjectData Projects, ProjectKeyArray


End Sub