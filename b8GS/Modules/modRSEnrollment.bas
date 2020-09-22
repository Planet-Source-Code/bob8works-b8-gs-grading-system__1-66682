Attribute VB_Name = "modRSEnrollment"
Option Explicit

Public Type tEnrollment
    
    EnrollmentID As String
    FK_StudentID As Long
    FK_SYID As Integer
    FK_CourseID As Integer
    YL As Integer
    Sem As Byte
    CreationDate As Date
    ModifiedDate As Date
    Remarks As String
    
End Type


Public CurrentEnrollment As tEnrollment

Public Function AddEnrollment(vEnrollment As tEnrollment) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddEnrollment = False
    
    sSQL = "Select * " & _
            " FROM tblEnrollment" & _
            " WHERE EnrollmentID='" & vEnrollment.EnrollmentID & "'"
    

    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modEnrollment", "AddEnrollment", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid Enrollment ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteEnrollment(vRS, vEnrollment) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddEnrollment = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditEnrollment(vEnrollment As tEnrollment) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditEnrollment = False
    
    sSQL = "Select * " & _
            " FROM tblEnrollment" & _
            " WHERE EnrollmentID='" & vEnrollment.EnrollmentID & "'"

    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modEnrollment", "EditEnrollment", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid Enrollment ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteEnrollment(vRS, vEnrollment) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditEnrollment = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteEnrollment(ByVal sEnrollmentID As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteEnrollment = False
    
    sSQL = "DELETE * " & _
            " FROM tblEnrollment" & _
            " WHERE EnrollmentID='" & sEnrollmentID & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modEnrollment", "DeleteEnrollment", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteEnrollment = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetEnrollmentByID(ByVal sEnrollmentID As String, vEnrollment As tEnrollment) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetEnrollmentByID = False
    
    sSQL = "SELECT * " & _
            " FROM tblEnrollment" & _
            " WHERE EnrollmentID='" & sEnrollmentID & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modEnrollment", "GetEnrollmentByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadEnrollment(vRS, vEnrollment) = False Then
        GoTo RAE
    End If
    
    GetEnrollmentByID = True
    
RAE:
    Set vRS = Nothing
End Function




Public Function AnyEnrollmentExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyEnrollmentExist = False
    
    sSQL = "SELECT * FROM tblEnrollment"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modEnrollment", "AnyEnrollmentExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyEnrollmentExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadEnrollment(ByRef vRS As ADODB.Recordset, ByRef vEnrollment As tEnrollment) As Boolean
    
    'default
    ReadEnrollment = False
    
    On Error GoTo RAE
    
    With vEnrollment
        .EnrollmentID = ReadField(vRS.Fields("EnrollmentID"))
        .FK_StudentID = ReadField(vRS.Fields("FK_StudentID"))
        .FK_SYID = ReadField(vRS.Fields("FK_SYID"))
        .FK_CourseID = ReadField(vRS.Fields("FK_CourseID"))
        .YL = ReadField(vRS.Fields("YL"))
        .Sem = ReadField(vRS.Fields("Sem"))
        .CreationDate = ReadField(vRS.Fields("CreationDate"))
        .ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
        .Remarks = ReadField(vRS.Fields("Remarks"))
    End With
    
    ReadEnrollment = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteEnrollment(ByRef vRS As ADODB.Recordset, ByRef vEnrollment As tEnrollment) As Boolean
    
    'default
    WriteEnrollment = False
    
    On Error GoTo RAE
    
    With vEnrollment
        vRS.Fields("EnrollmentID") = .EnrollmentID
        vRS.Fields("FK_StudentID") = .FK_StudentID
        vRS.Fields("FK_SYID") = .FK_SYID
        vRS.Fields("FK_CourseID") = .FK_CourseID
        vRS.Fields("YL") = .YL
        vRS.Fields("Sem") = .Sem
        vRS.Fields("CreationDate") = .CreationDate
        vRS.Fields("ModifiedDate") = .ModifiedDate
        vRS.Fields("Remarks") = .Remarks
    End With
    
    WriteEnrollment = True
    Exit Function
    
RAE:
    
End Function



