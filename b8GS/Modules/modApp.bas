Attribute VB_Name = "modApp"
Option Explicit

Public Function WriteErrorLog(sModuleName As String, sRoutineName As String, sDetail As String)
    frmErrMsg.ShowForm sModuleName, sRoutineName, sDetail
End Function
