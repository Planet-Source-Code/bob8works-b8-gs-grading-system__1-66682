Attribute VB_Name = "modRSSY"
Option Explicit

Public Type tSY
    
    SYID As Integer
    SYTitle As String
    
End Type



Public Function AddSY(vSY As tSY) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddSY = False
    
    sSQL = "SELECT * FROM tblSY WHERE SYID=" & vSY.SYID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSY", "AddSY", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid SY ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteSY(vRS, vSY) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddSY = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditSY(vSY As tSY) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditSY = False
    
    sSQL = "SELECT * FROM tblSY WHERE SYID=" & vSY.SYID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSY", "EditSY", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid SY ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteSY(vRS, vSY) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditSY = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteSY(iSYID As Integer) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteSY = False
    
    sSQL = "DELETE * FROM tblSY WHERE SYID=" & iSYID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSY", "DeleteSY", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteSY = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSYByID(iSYID As Integer, vSY As tSY) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSYByID = False
    
    sSQL = "SELECT * FROM tblSY WHERE SYID=" & iSYID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSY", "GetSYByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSY(vRS, vSY) = False Then
        GoTo RAE
    End If
    
    GetSYByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSYBySYTitle(sSYTitle As String, vSY As tSY) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSYBySYTitle = False
    
    sSQL = "SELECT * FROM tblSY WHERE SYTitle='" & sSYTitle & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSY", "GetSYBySYTitle", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSY(vRS, vSY) = False Then
        GoTo RAE
    End If
    
    GetSYBySYTitle = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSYTitle(ByVal iSYID As Integer) As String
    
    GetSYTitle = iSYID & "-" & iSYID + 1
    
End Function


Public Function GetSYID(ByVal sSYTitle As String) As Integer
    
    'default
    GetSYID = -1
    
    If Len(Trim(sSYTitle)) < 4 Then
        Exit Function
    End If
    
    If IsNumeric(Left(Trim(sSYTitle), 4)) = False Then
        Exit Function
    End If
    
    GetSYID = Val(Left(Trim(sSYTitle), 4))
    
End Function



Public Function AnySYExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnySYExist = False
    
    sSQL = "SELECT * FROM tblSY"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSY", "AnySYExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnySYExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadSY(ByRef vRS As ADODB.Recordset, ByRef vSY As tSY) As Boolean
    
    'default
    ReadSY = False
    
    On Error GoTo RAE
    
    With vSY
        .SYID = ReadField(vRS.Fields("SYID"))
        .SYTitle = ReadField(vRS.Fields("SYTitle"))
    End With
    
    ReadSY = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteSY(ByRef vRS As ADODB.Recordset, ByRef vSY As tSY) As Boolean
    
    'default
    WriteSY = False
    
    On Error GoTo RAE
    
    With vSY
        vRS.Fields("SYID") = .SYID
        vRS.Fields("SYTitle") = .SYTitle
    End With
    
    WriteSY = True
    Exit Function
    
RAE:
    
End Function



