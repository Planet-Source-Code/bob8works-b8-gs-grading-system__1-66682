Attribute VB_Name = "modRSStudent"
Option Explicit

Public Type tStudent
    
    StudentID As Long
    FirstName As String
    MiddleName As String
    LastName As String
    YL As Integer
    CreationDate As Date
    ModifiedDate As Date
    Active As Boolean
    
End Type


Public CurrentStudent As tStudent

Public Function AddStudent(vStudent As tStudent) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddStudent = False
    
    sSQL = "SELECT * FROM tblStudent WHERE StudentID=" & vStudent.StudentID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudent", "AddStudent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid Student ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteStudent(vRS, vStudent) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddStudent = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditStudent(vStudent As tStudent) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditStudent = False
    
    sSQL = "SELECT * FROM tblStudent WHERE StudentID=" & vStudent.StudentID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudent", "EditStudent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid Student ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteStudent(vRS, vStudent) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditStudent = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function UpdateStudentYL(ByVal iStudentID As Long, ByVal iYL As Integer) As Boolean
    
    Dim vStudent As tStudent
   
    'default
    UpdateStudentYL = False
    
    If GetStudentByID(iStudentID, vStudent) = False Then
        GoTo RAE
    End If
    
    If iYL > vStudent.YL Then
        vStudent.YL = iYL
        
        If EditStudent(vStudent) = False Then
            GoTo RAE
        End If
    End If
    
    'return success
    UpdateStudentYL = True
    
RAE:
End Function

Public Function DeleteStudent(iStudentID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteStudent = False
    
    sSQL = "DELETE * FROM tblStudent WHERE StudentID=" & iStudentID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudent", "DeleteStudent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteStudent = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetStudentByID(iStudentID As Long, vStudent As tStudent) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetStudentByID = False
    
    sSQL = "SELECT * FROM tblStudent WHERE StudentID=" & iStudentID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudent", "GetStudentByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadStudent(vRS, vStudent) = False Then
        GoTo RAE
    End If
    
    GetStudentByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyStudentExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyStudentExist = False
    
    sSQL = "SELECT * FROM tblStudent"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modStudent", "AnyStudentExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyStudentExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadStudent(ByRef vRS As ADODB.Recordset, ByRef vStudent As tStudent) As Boolean
    
    'default
    ReadStudent = False
    
    On Error GoTo RAE
    
    With vStudent
        .StudentID = vRS.Fields("StudentID")
        .StudentID = vRS.Fields("StudentID")
        .FirstName = vRS.Fields("FirstName")
        .MiddleName = vRS.Fields("MiddleName")
        .LastName = vRS.Fields("LastName")
        .YL = vRS.Fields("YL")
        .CreationDate = vRS.Fields("CreationDate")
        .ModifiedDate = vRS.Fields("ModifiedDate")
       .Active = vRS.Fields("Active")
    End With
    
    ReadStudent = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteStudent(ByRef vRS As ADODB.Recordset, ByRef vStudent As tStudent) As Boolean
    
    'default
    WriteStudent = False
    
    On Error GoTo RAE
    
    With vStudent
        vRS.Fields("StudentID") = .StudentID
        vRS.Fields("StudentID") = .StudentID
        vRS.Fields("FirstName") = .FirstName
        vRS.Fields("MiddleName") = .MiddleName
        vRS.Fields("LastName") = .LastName
        vRS.Fields("YL") = .YL
        vRS.Fields("CreationDate") = .CreationDate
        vRS.Fields("ModifiedDate") = .ModifiedDate
        vRS.Fields("Active") = .Active
    End With
    
    WriteStudent = True
    Exit Function
    
RAE:
    
End Function



