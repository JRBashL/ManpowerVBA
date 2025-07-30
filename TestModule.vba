' Test
Public team As TeamMembers
Public TestProject As ProjectBlockClass

Sub CreateTest()
    Set team = New TeamMembers
    team.Constructor
    
    Dim blockheight As Integer
    blockheight = team.TeamMembersDict.Count
    
    Set TestProject = New ProjectBlockClass
    
    TestProject.Constructor "My Cool Project", 123456, 2, blockheight + 1, 35
    TestProject.AddProjectBlock team
End Sub