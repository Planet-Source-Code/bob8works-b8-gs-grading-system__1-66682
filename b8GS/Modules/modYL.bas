Attribute VB_Name = "modYL"
Option Explicit

Public Function YLNumToStr(ByVal iYL As Integer) As String
    Select Case iYL
        Case 1
            YLNumToStr = "I"
        Case 2
            YLNumToStr = "II"
        Case 3
            YLNumToStr = "III"
        Case 4
            YLNumToStr = "IV"
        Case 5
            YLNumToStr = "V"
        Case 6
            YLNumToStr = "VI"
        Case 7
            YLNumToStr = "VII"
        Case 8
            YLNumToStr = "VIII"
        Case 9
            YLNumToStr = "IX"
        Case 10
            YLNumToStr = "X"
        Case Else
            YLNumToStr = "?"
    End Select
End Function



Public Function YLStrToNum(ByVal sYL As String) As Integer
    Select Case sYL
        Case "I"
            YLStrToNum = 1
        Case "II"
            YLStrToNum = 2
        Case "III"
            YLStrToNum = 3
        Case "IV"
            YLStrToNum = 4
        Case "V"
            YLStrToNum = 5
        Case "VI"
            YLStrToNum = 6
        Case "VII"
            YLStrToNum = 7
        Case "VIII"
            YLStrToNum = 8
        Case "IX"
            YLStrToNum = 9
        Case "X"
            YLStrToNum = 10
        Case Else
            YLStrToNum = 0
    End Select
End Function
