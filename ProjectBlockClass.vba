' === ProjectBlock.vba ===
' This is in a class module
' Strong type things like in C#
Option Explicit

'--- Properties ---
Private v_projectName As String
Private v_projectNumber As Variant
Private v_headRow As Integer
Private v_blockHeight As Integer
Private v_blockLength As Integer

'--- Constructor-like method ---
Public Sub Constructor(ByVal projectName As String, ByVal projectNumber as String, ByVal headRow as Integer, ByVal blockHeight As Integer, ByVal blockLength As Integer)
    v_projectName = projectname
    v_projectNumber = projectNumber
    v_headRow = headRow
    v_blockHeight = blockHeight
    v_blockLength = blockLength
End Sub

'--- Getters / Setters ---
Public Property Get ProjectName() As String
    ProjectName = v_projectName
End Property

Public Property Let ProjectName(ByVal value As String)
    v_projectName = ProjectName
End Property

Public Property Get ProjectNumber() As Variant
    ProjectNumber = v_projectNumber
End Property

Public Property Let ProjectNumber(ByVal value as Variant)
    v_projectNumber = value
End Property

Public Property Get HeadRow() As Integer
    HeadRow = v_headRow
End Property

Public Property Let HeadRow(ByVal value As Integer)
    v_headRow = value
End Property

'--- Example Method ---
Public Sub AddProjectBlock(team As TeamMembers)
    ' Shorthand worksheets
    Dim ws As Worksheet
    Set ws = Worksheets("Alberta")

    ' Insert Rows
    Dim i as Integer
    For i = 1 to v_blockHeight 
        ws.Rows(v_headRow).Insert Shift:=xlDown
        Debug.Print v_headRow
    Next i

    ' Populate Project Name and Project Number
    ws.Cells(v_headRow, "A").Value = v_projectName
    ws.Cells(v_headRow + 2, 1).Value = v_projectNumber

    ' Populate Team 
    ws.Cells(v_headRow, 2).Value = "*"
    For i = 1 to v_blockHeight - 1
        ws.Cells(v_headRow + i, 2).Value = team.TeamMembersNum(i)
    Next i

    ' Formatting
    Worksheets("Template").Range("A1:W21").Copy
    ws.Range("A" & v_headRow & ":W" & (v_headRow + v_blockHeight - 1 )).PasteSpecial xlPasteFormats
    ws.columns(1).ColumnWidth = 64.14
    ws.columns(2).ColumnWidth = 11
    for i = 3 to v_blockLength
        ws.columns(i).ColumnWidth = 10
    Next i
    for i = v_headRow to v_headRow + v_blockHeight
        ws.rows(i).RowHeight = 15
    Next i
    

End Sub


