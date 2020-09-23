Attribute VB_Name = "ModProcedures"
Option Explicit


Public Sub SQL_DELETE(ByVal strSQL As String)
Set COMMAND_DELETE = New ADODB.Command
    With COMMAND_DELETE
        .ActiveConnection = conn
        .CommandText = strSQL
        .Execute
    End With
End Sub
Public Sub SQL_INSERT(ByVal strSQL As String)
Set COMMAND_INSERT = New ADODB.Command
    With COMMAND_INSERT
        .ActiveConnection = conn
        .CommandText = strSQL
        .Execute
    End With
End Sub

Public Sub SQL_UPDATE(ByVal strSQL As String)
Set COMMAND_UPDATE = New ADODB.Command
    With COMMAND_UPDATE
        .ActiveConnection = conn
        .CommandText = strSQL
        .Execute
    End With
End Sub
Public Sub LogIn()
    Set LogInRS = New ADODB.Recordset
    With LogInRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM TIMELOG"
    End With
End Sub

Public Sub Info()
    Set InfoRS = New ADODB.Recordset
    With InfoRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM INFO"
    End With
End Sub
Public Sub Sum()
    Set SumRS = New ADODB.Recordset
    With SumRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM SUMMARY"
    End With
End Sub
Public Sub Sec()
    Set SecRS = New ADODB.Recordset
    With SecRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM SECURITY"
    End With
End Sub

Public Sub DateRange()
    Set DateRangeRS = New ADODB.Recordset
    With DateRangeRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM DATERANGE"
    End With
End Sub

Public Sub CompanySettings()
    Set CompanyRS = New ADODB.Recordset
    With CompanyRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM COMPANY_INFO"
    End With
End Sub


