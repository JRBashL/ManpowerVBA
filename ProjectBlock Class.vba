' === ProjectBlock.cls ===

Option Explicit

'--- Properties ---
Private _projectName As String
Private pStartDate As Date
Private pEndDate As Date
Private pRowStart As Long
Private pRowEnd As Long

'--- Constructor-like method ---
Public Sub Constructor(ByVal name As String, ByVal startRow As Long, ByVal endRow As Long)
    pName = name
    pRowStart = startRow
    pRowEnd = endRow
    ' Optional: parse dates from sheet if needed
End Sub

'--- Getters / Setters ---
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal value As String)
    pName = value
End Property

Public Property Get RowStart() As Long
    RowStart = pRowStart
End Property

Public Property Get RowEnd() As Long
    RowEnd = pRowEnd
End Property

'--- Example Method ---
Public Function DurationInDays() As Long
    DurationInDays = pEndDate - pStartDate
End Function

Public Sub LoadDates(ByVal ws As Worksheet)
    '
