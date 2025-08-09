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
    i = 2 ' =2 due to the title being row 1 in Worksheets("Scripting")
    set ws = Worksheets("Scripting")

    ' Goes through the data in the worksheet and if a gap in both columns are found, it ends
    Do While endoflist = False
        if IsEmpty(ws.Cells(i,4)) And IsEmpty(ws.Cells(i, 5)) Then 
            Exit Do
        else
            TeamMembersNum.Add ws.Cells(i, 4).Value, ws.Cells(i, 5).Value
            TeamMembersName.Add ws.Cells(i, 5).Value, ws.Cells(i, 4).Value
            i = i + 1           
        end if
    Loop

End Sub


