Attribute VB_Name = "modFactory"
Option Explicit

' We don't have static methods/constructors so we use a factory module for this purpose
Public Function CreatePerson(ByVal FirstName As String, ByVal LastName As String, ByVal Yob As Integer) As clsPerson
    Set CreatePerson = New clsPerson
    CreatePerson.FirstName = FirstName
    CreatePerson.LastName = LastName
    CreatePerson.Yob = Yob
End Function

