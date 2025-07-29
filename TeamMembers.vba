' === ProjectBlock.vba ===

' Strong type things like in C#
Option Explicit

'--- Properties ---
Public TeamMembers As New Collection
Public TeamSize As Integer

Public Sub CreateTeamCollection()
    TeamMembers.Add("Pertti")
    TeamMembers.Add("Martin")
    TeamMembers.Add("EIT-Edm")
    TeamMembers.Add("Craig")
    TeamMembers.Add("Mike T")
    TeamMembers.Add("Jiaxun")
    TeamMembers.Add("EIT-Calgary")
    TeamMembers.Add("Mau")
    TeamMembers.Add("Quinn")
    TeamMembers.Add("Jorrell")
    TeamMembers.Add("Syed")
    TeamMembers.Add("Ping")
    TeamMembers.Add("Yang")
    TeamMembers.Add("Denis")
    TeamMembers.Add("Artem")
    TeamMembers.Add("Christy")
    TeamMembers.Add("EC -Marian")
    TeamMembers.Add("Moosa")
    TeamMembers.Add("Shyam")
    TeamMembers.Add("(EIT/Others)")
End Sub

TeamSize = TeamMembers.Count
