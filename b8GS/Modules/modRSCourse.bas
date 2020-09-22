Attribute VB_Name = "modRSCourse"
Option Explicit

Public Type tCourse
    
    CourseID As Integer
    Title As String
    Description As String
    
End Type


Public CurrentCourse As tCourse

Public Function AddCourse(vCourse As tCourse) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddCourse = False
    
    sSQL = "SELECT * FROM tblCourse WHERE CourseID=" & vCourse.CourseID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "AddCourse", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid Course ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteCourse(vRS, vCourse) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddCourse = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditCourse(vCourse As tCourse) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditCourse = False
    
    sSQL = "SELECT * FROM tblCourse WHERE CourseID=" & vCourse.CourseID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "EditCourse", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid Course ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteCourse(vRS, vCourse) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditCourse = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteCourse(iCourseID As Integer) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteCourse = False
    
    sSQL = "DELETE * FROM tblCourse WHERE CourseID=" & iCourseID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "DeleteCourse", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteCourse = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetCourseByID(iCourseID As Integer, vCourse As tCourse) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCourseByID = False
    
    sSQL = "SELECT * FROM tblCourse WHERE CourseID=" & iCourseID
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "GetCourseByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCourse(vRS, vCourse) = False Then
        GoTo RAE
    End If
    
    GetCourseByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetCourseByTitle(sTitle As String, vCourse As tCourse) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCourseByTitle = False
    
    sSQL = "SELECT * FROM tblCourse WHERE Title='" & sTitle & "'"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "GetCourseByTitle", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCourse(vRS, vCourse) = False Then
        GoTo RAE
    End If
    
    GetCourseByTitle = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function GetNewCourseID() As Integer
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewCourseID = 0
    
    sSQL = "SELECT Max(tblCourse.CourseID)+1 AS MaxOfCourseID" & _
            " FROM tblCourse;"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "GetNewCourseID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid Course ID. It does not exist in record.", vbExclamation
        GetNewCourseID = 1
        GoTo RAE
    End If
    
    On Error GoTo RAE
    GetNewCourseID = 1
    GetNewCourseID = ReadField(vRS.Fields("MaxOfCourseID"))
    
RAE:
    Set vRS = Nothing
End Function


Public Function AnyCourseExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyCourseExist = False
    
    sSQL = "SELECT * FROM tblCourse"
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog "modCourse", "AnyCourseExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyCourseExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadCourse(ByRef vRS As ADODB.Recordset, ByRef vCourse As tCourse) As Boolean
    
    'default
    ReadCourse = False
    
    On Error GoTo RAE
    
    With vCourse
        .CourseID = ReadField(vRS.Fields("CourseID"))
        .Title = ReadField(vRS.Fields("Title"))
        .Description = ReadField(vRS.Fields("Description"))
    End With
    
    ReadCourse = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteCourse(ByRef vRS As ADODB.Recordset, ByRef vCourse As tCourse) As Boolean
    
    'default
    WriteCourse = False
    
    On Error GoTo RAE
    
    With vCourse
        vRS.Fields("CourseID") = .CourseID
        vRS.Fields("Title") = .Title
        vRS.Fields("Description") = .Description
    End With
    
    WriteCourse = True
    Exit Function
    
RAE:
    
End Function

