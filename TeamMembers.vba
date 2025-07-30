' === TeamMembers.vba ===
' This is in a class module

' Strong type things like in C#
Option Explicit

'--- Properties ---
Public TeamMembersDict As Object
Public TeamSize As Integer

'--- Constructor-like method ---
Public Sub Constructor()
    Set TeamMembersDict = CreateObject("Scripting.Dictionary")
    CreateTeamCollection
    TeamSize = TeamMembersDict.Count
End Sub

Private Sub CreateTeamCollection()

    TeamMembersDict.RemoveAll

    dim endoflist As Boolean
    dim i As Integer
    dim ws As Worksheet

    endoflist = false
    i = 1
    set ws = WorkSheets("Team")

    ' Goes through the data in the worksheet and if a gap in both columns are found, it ends
    Do While endoflist = false
        if IsEmpty(ws.Cells(i,1)) And IsEmpty(ws.Cells(i, 2)) Then 
            Exit Do
        else
            TeamMembersDict.Add ws.Cells(i, 1).Value, ws.Cells(i, 2).Value
            i = i + 1
            
        end if
    Loop

End Sub


