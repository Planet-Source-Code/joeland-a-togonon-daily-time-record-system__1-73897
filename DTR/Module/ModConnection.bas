Attribute VB_Name = "ModConnection"
Option Explicit


Public Sub DBLOAD()

    Set conn = New ADODB.Connection
        
        With conn
            .Provider = "Microsoft.jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\database\log.mdb;Persist Security Info= False"
            .Open
        End With
        
   frmDTR.Timer1.Enabled = True
End Sub


Public Function ConRs(ByRef smlaDb As ADODB.Connection, ByRef smlars As ADODB.Recordset, smlaSql As String) As Boolean


Set smlars = Nothing
Set smlars = New ADODB.Recordset

smlars.Open smlaSql, smlaDb, adOpenStatic, adLockOptimistic
ConRs = True

End Function
