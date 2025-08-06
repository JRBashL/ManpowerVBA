' === TeamMembers.vba ===
' This is in a class module

' Strong type things like in C#
Option Explicit

'--- Properties ---
Public TeamMembersNum As Object
Public TeamMembersName As Object
Public TeamSize As Integer

'--- Constructor-like method ---
Public Sub Constructor()
    Set TeamMembersNum = CreateObject("Scripting.Dictionary")
    Set TeamMembersName = CreateObject("Scripting.Dictionary")
    CreateTeamCollection
    TeamSize = TeamMembersNum.Count
End Sub

Private Sub CreateTeamCollection()

    If Not TeamMembersNum Is Nothing Then TeamMembersNum.RemoveAll
    If Not TeamMembersName Is Nothing Then TeamMembersName.RemoveAll

    Dim endoflist As Boolean
    Dim i As Integer
    Dim ws As Worksheet

    endoflist = False
    i = 1
    set ws = Worksheets("Team")

    ' Goes through the data in the worksheet and if a gap in both columns are found, it ends
    Do While endoflist = False
        if IsEmpty(ws.Cells(i,1)) And IsEmpty(ws.Cells(i, 2)) Then 
            Exit Do
        else
            TeamMembersNum.Add ws.Cells(i, 1).Value, ws.Cells(i, 2).Value
            TeamMembersName.Add ws.Cells(i, 2).Value, ws.Cells(i, 1).Value
            i = i + 1           
        end if
    Loop

End Sub


