' === ProjectBlock.vba ===

' Strong type things like in C#
Option Explicit

'--- Properties ---
Private _projectName As String
Private _projectNumber As Integer
Private _headRow As Integer

'--- Constructor-like method ---
Public Sub Constructor(ByVal projectName As String, ByVal projectNumber as Integer, ByVal headRow as String)
    _projectName = name
    _projectNumber = projectNumber
    _headRow = headRow
    ' Optional: parse dates from sheet if needed
End Sub

'--- Getters / Setters ---
Public Property Get ProjectName() As String
    ProjectName = _projectName
End Property

Public Property Let ProjectName(ByVal value As String)
    _projectName = ProjectName
End Property

Public Property Get ProjectNumber() As Integer
    ProjectNumber = _projectNumber
End Property

Public Property Let ProjectNumber(ByVal value as Integer)
    _projectNumber = value
End Property

Public Property Get HeadRow() As String
    HeadRow = _headRow
End Property

Public Property Let ProjectNumber(ByVal value as String)
    _projectNumber = value
End Property

'--- Example Method ---
Public Sub AddProjectBlock()
    Dim ws As Worksheets
    Set ws = Worksheets("Alberta")

    ws.Cells(_headRow, 1).Value = _projectName
    ws("Sheet1").Range(_headRow + 2, 1).Value = _projectNumber


   
End Sub


