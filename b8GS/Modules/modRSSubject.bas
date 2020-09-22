Attribute VB_Name = "modRSSubject"
Option Explicit

Public Type tSubject
    
    SubjectID As String
    FK_SYID As Integer
    Sem As Byte
    Term As Byte
    CN As Integer
    
    SubjectTitle  As String
    SubjectAbr As String
    PreReqCN As Integer
    
    TimeIn As Integer
    TimeOut As Integer
    Room As String
    SchedDay As String
    Instructor As String
    
End Type


Public CurrentSubject As tSubject

Public Function AddSubject(vSubject As tSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddSubject = False
    
    sSQL = "SELECT * FROM tblSubject" & _
            " WHERE SubjectID='" & vSubject.SubjectID & "'"
    
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "AddSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid Subject ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    On Error GoTo RAE
    'add new record
    vRS.AddNew
    
    If WriteSubject(vRS, vSubject) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddSubject = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditSubject(vSubject As tSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditSubject = False
    
    sSQL = "SELECT * FROM tblSubject" & _
            " WHERE SubjectID='" & vSubject.SubjectID & "'"

    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "EditSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid Subject ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteSubject(vRS, vSubject) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditSubject = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteSubject(ByVal sSubjectID As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteSubject = False
    
    sSQL = "DELETE * FROM tblSubject WHERE SubjectID='" & sSubjectID & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "DeleteSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteSubject = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSubjectByID(ByVal sSubjectID As String, ByRef vSubject As tSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSubjectByID = False
    
    sSQL = "SELECT * FROM tblSubject" & _
            " WHERE SubjectID='" & sSubjectID & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "GetSubjectByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    If ReadSubject(vRS, vSubject) = False Then
        GoTo RAE
    End If
    
    
    GetSubjectByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSubjectBySubjectTitle(sSubjectTitle As String, vSubject As tSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSubjectBySubjectTitle = False
    
    sSQL = "SELECT * FROM tblSubject WHERE SubjectTitle='" & sSubjectTitle & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "GetSubjectBySubjectTitle", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSubject(vRS, vSubject) = False Then
        GoTo RAE
    End If
    
    GetSubjectBySubjectTitle = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSubjectBySubjectAbr(sSubjectAbr As String, vSubject As tSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSubjectBySubjectAbr = False
    
    sSQL = "SELECT * FROM tblSubject WHERE SubjectAbr='" & sSubjectAbr & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "GetSubjectBySubjectAbr", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSubject(vRS, vSubject) = False Then
        GoTo RAE
    End If
    
    GetSubjectBySubjectAbr = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function AnySubjectExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnySubjectExist = False
    
    sSQL = "SELECT * FROM tblSubject"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modSubject", "AnySubjectExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnySubjectExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadSubject(ByRef vRS As ADODB.Recordset, ByRef vSubject As tSubject) As Boolean
    
    'default
    ReadSubject = False
    
    On Error GoTo RAE
    
    With vSubject
    
        
        .SubjectID = vRS.Fields("SubjectID")
        .FK_SYID = ReadField(vRS.Fields("FK_SYID"))
        .CN = ReadField(vRS.Fields("CN"))
        .Sem = ReadField(vRS.Fields("Sem"))
        .Term = ReadField(vRS.Fields("Term"))
        
        .SubjectTitle = ReadField(vRS.Fields("SubjectTitle"))
        .SubjectAbr = ReadField(vRS.Fields("SubjectAbr"))
        .PreReqCN = ReadField(vRS.Fields("PreReqCN"))
        
        .TimeIn = ReadField(vRS.Fields("TimeIn"))
        .TimeOut = ReadField(vRS.Fields("TimeOut"))
        .Room = ReadField(vRS.Fields("Room"))
        .SchedDay = ReadField(vRS.Fields("SchedDay"))
        
        .Instructor = ReadField(vRS.Fields("Instructor"))
        
    End With
    
    ReadSubject = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteSubject(ByRef vRS As ADODB.Recordset, ByRef vSubject As tSubject) As Boolean
    
    'default
    WriteSubject = False
    
    On Error GoTo RAE
    
    With vSubject
        
        vRS.Fields("SubjectID") = .SubjectID
        vRS.Fields("FK_SYID") = .FK_SYID
        vRS.Fields("Sem") = .Sem
        vRS.Fields("Term") = .Term
        vRS.Fields("CN") = .CN
        
        vRS.Fields("SubjectTitle") = .SubjectTitle
        vRS.Fields("SubjectAbr") = .SubjectAbr
        vRS.Fields("PreReqCN") = .PreReqCN
        vRS.Fields("TimeIn") = .TimeIn
        vRS.Fields("TimeOut") = .TimeOut
        vRS.Fields("Room") = .Room
        vRS.Fields("SchedDay") = .SchedDay
        
        vRS.Fields("Instructor") = .Instructor
    
    End With
    
    WriteSubject = True
    Exit Function
    
RAE:
    
End Function



