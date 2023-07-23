Attribute VB_Name = "modRandom"
Option Explicit

Public Function RandInt(ByVal min As Integer, ByVal max As Integer)
    ' Click inside a word an press F1 to open the documentation
    Call Randomize

    RandInt = Int((max - min + 1) * Rnd() + min)
End Function

