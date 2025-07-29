' === ProjectBlock.vba ===
' This is in a class module
' Strong type things like in C#
Option Explicit

'--- Properties ---
Private v_projectName As String
Private v_projectNumber As String
Private v_headRow As Integer
Private v_blockHeight As Integer

'--- Constructor-like method ---
Public Sub Constructor(ByVal projectName As String, ByVal projectNumber as String, ByVal headRow as Integer)
    v_projectName = projectname
    v_projectNumber = projectNumber
    v_headRow = headRow
    v_blockHeight = TeamSize + 1
End Sub

'--- Getters / Setters ---
Public Property Get ProjectName() As String
    ProjectName = v_projectName
End Property

Public Property Let ProjectName(ByVal value As String)
    v_projectName = ProjectName
End Property

Public Property Get ProjectNumber() As String
    ProjectNumber = v_projectNumber
End Property

Public Property Let ProjectNumber(ByVal value as String)
    v_projectNumber = value
End Property

Public Property Get HeadRow() As Integer
    HeadRow = v_headRow
End Property

Public Property Let HeadRow(ByVal value As Integer)
    v_headRow = value
End Property

'--- Example Method ---
Public Sub AddProjectBlock()
    ' Shorthand worksheets
    Dim ws As Worksheet
    Set ws = Worksheets("Alberta")

    ' Populate the rest of the column as blanks
    Dim i as Integer
    For i = 1 to v_blockHeight 
        ws.Rows(v_headRow).Insert Shift:=xlDown
        ws.Rows(v_headRow + i - 1).Interior.Color = RGB(0, 0, 255)
        Debug.Print v_headRow
    Next i
    ' Populate Project Name and Project Number
    ws.Cells(v_headRow, "A").Value = v_projectName
    ws.Cells(v_headRow + 2, 1).Value = v_projectNumber


   
End Sub


