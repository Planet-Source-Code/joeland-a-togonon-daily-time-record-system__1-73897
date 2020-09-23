Attribute VB_Name = "Module1"
Option Explicit


Public Function CONNECT_TO_ACCESS() As Boolean
    Dim isOPEN As Boolean
    Dim REPLY As VbMsgBoxResult
          
    strServer = "127.0.0.1"

    isOPEN = False
    On Error GoTo ERR_CONNECT
    Do Until isOPEN = True
        Set conn = New ADODB.Connection
        conn.CursorLocation = adUseClient
        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\log.mdb;Persist Security Info=False;Jet OLEDB:Database Password=;"
        conn.Open
        
        isOPEN = True
        Loop
        CONNECT_TO_ACCESS = isOPEN
        Exit Function
    
ERR_CONNECT:
        REPLY = MsgBox("Error Number:" & err.Number & vbNewLine & "Description:" & err.Description, vbExclamation + vbRetryCancel, "ERROR CONNECTION")
        If REPLY = vbCancel Then
            CONNECT_TO_ACCESS = False
        ElseIf REPLY = vbRetry Then
            Resume
           ' frmDTR.Timer1.Enabled = True
        End If
        
End Function



Public Sub SQLDB(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\log.mdb;Persist Security Info=False; Jet OLEDB:Database Password ="
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB1(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\log.mdb;Persist Security Info=False; Jet OLEDB:Database Password ="
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB2(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\log.mdb;Persist Security Info=False; Jet OLEDB:Database Password ="
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub

