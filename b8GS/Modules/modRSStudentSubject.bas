Attribute VB_Name = "modRSStudentSubject"
Option Explicit

Public Type tStudentSubject
    
    FK_EnrollmentID As String
    FK_SubjectID As String
    Grade As Double
    CreationDate As Date
    ModifiedDate As Date
    Remarks As String
    
End Type


Public CurrentStudentSubject As tStudentSubject

Public Function AddStudentSubject(vStudentSubject As tStudentSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddStudentSubject = False
    
    sSQL = "Select * " & _
            " FROM tblStudentSubject" & _
            " WHERE FK_EnrollmentID='" & vStudentSubject.FK_EnrollmentID & "' AND FK_SubjectID='" & vStudentSubject.FK_SubjectID & "'"
    

    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudentSubject", "AddStudentSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid StudentSubject ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteStudentSubject(vRS, vStudentSubject) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddStudentSubject = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditStudentSubject(vStudentSubject As tStudentSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditStudentSubject = False
    
    sSQL = "Select * " & _
            " FROM tblStudentSubject" & _
            " WHERE FK_EnrollmentID='" & vStudentSubject.FK_EnrollmentID & "' AND FK_SubjectID='" & vStudentSubject.FK_SubjectID & "'"

    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudentSubject", "EditStudentSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid StudentSubject ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteStudentSubject(vRS, vStudentSubject) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditStudentSubject = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteStudentSubject(ByVal sEnrollmentID As String, ByVal sSubjectID As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteStudentSubject = False
    
    sSQL = "DELETE * " & _
            " FROM tblStudentSubject" & _
            " WHERE FK_EnrollmentID='" & sEnrollmentID & "' AND FK_SubjectID='" & sSubjectID & "'"
            
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudentSubject", "DeleteStudentSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteStudentSubject = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetStudentSubjectByID(ByRef EnrollmentID As String, ByVal SubjectID As String, vStudentSubject As tStudentSubject) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetStudentSubjectByID = False
    
    sSQL = "SELECT * " & _
            " FROM tblStudentSubject" & _
            " WHERE FK_EnrollmentID='" & EnrollmentID & "' AND FK_SubjectID'=" & SubjectID & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudentSubject", "GetStudentSubjectByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadStudentSubject(vRS, vStudentSubject) = False Then
        GoTo RAE
    End If
    
    GetStudentSubjectByID = True
    
RAE:
    Set vRS = Nothing
End Function




Public Function AnyStudentSubjectExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyStudentSubjectExist = False
    
    sSQL = "SELECT * FROM tblStudentSubject"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudentSubject", "AnyStudentSubjectExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyStudentSubjectExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadStudentSubject(ByRef vRS As ADODB.Recordset, ByRef vStudentSubject As tStudentSubject) As Boolean
    
    'default
    ReadStudentSubject = False
    
    On Error GoTo RAE
    
    With vStudentSubject
        
        .FK_EnrollmentID = vRS.Fields("FK_EnrollmentID")
        .FK_SubjectID = vRS.Fields("FK_SubjectID")
        .Grade = vRS.Fields("Grade")
        .CreationDate = ReadField(vRS.Fields("CreationDate"))
        .ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
        .Remarks = ReadField(vRS.Fields("Remarks"))
    
    End With
    
    ReadStudentSubject = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteStudentSubject(ByRef vRS As ADODB.Recordset, ByRef vStudentSubject As tStudentSubject) As Boolean
    
    'default
    WriteStudentSubject = False
    
    On Error GoTo RAE
    
    With vStudentSubject

        vRS.Fields("FK_EnrollmentID") = .FK_EnrollmentID
        vRS.Fields("FK_SubjectID") = .FK_SubjectID
        vRS.Fields("Grade") = .Grade
        vRS.Fields("CreationDate") = .CreationDate
        vRS.Fields("ModifiedDate") = .ModifiedDate
        vRS.Fields("Remarks") = .Remarks
        
    End With
    
    WriteStudentSubject = True
    Exit Function
    
RAE:
    
End Function





