Attribute VB_Name = "modGV"
Option Explicit


Public GV_DateFormat As String

Public CurrentSY As tSY

Public Sub InitGV()
    
    GV_DateFormat = "dd-Mmm-yy"
    
    
    CurrentSY.SYID = Val(GetSetting(App.Title, "currentsetting", "currentsy", Year(Now)))
    CurrentSY.SYTitle = CurrentSY.SYID & "-" & CurrentSY.SYID + 1
End Sub

Public Function SetCureentSY(ByVal iSY As Integer) As Boolean
    
    On Error GoTo RAE
    'save current sy info
    SaveSetting App.Title, "currentsetting", "currentsy", iSY
    'refresh current sy info
    CurrentSY.SYID = Val(GetSetting(App.Title, "currentsetting", "currentsy", Year(Now)))
    CurrentSY.SYTitle = CurrentSY.SYID & "-" & CurrentSY.SYID + 1
    
    SetCureentSY = True
RAE:
End Function
