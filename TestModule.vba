' Create instance of the class

Sub CreateTestProjectBlock()

    CreateTeamCollection
    
    Dim TestProject As ProjectBlock
    Set TestProject = New ProjectBlock
    
    TestProject.Constructor "TestProject", "12345", 2
    TestProject.AddProjectBlock
End Sub
