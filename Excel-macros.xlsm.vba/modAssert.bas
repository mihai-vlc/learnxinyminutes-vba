Attribute VB_Name = "modAssert"
Option Explicit

' Use view -> call stack (Ctrl + L)
Public Sub Assert(ByVal condition As Boolean)
    Debug.Assert (condition)
End Sub

' Use view -> call stack (Ctrl + L)
Public Sub AssertCollection(c1 As Collection)
    Debug.Assert (Not c1 Is Nothing)
    Debug.Assert (c1.Count >= 1)
End Sub
