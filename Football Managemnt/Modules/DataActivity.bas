Attribute VB_Name = "DataActivity"
Option Explicit

'Dim connect As New ADODB.Connection


'Public Sub ConnectDb() 'Connect to Database

 '   On Error Resume Next
  '  connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb;Persist Security Info=False;Jet OLEDB:Database Password= reg"
    
'End Sub

Public Sub DataConn(adoObject As Adodc, adoRecord As String)

    On Error Resume Next
       
    'Loads the database
    adoObject.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb;Persist Security Info=False;Jet OLEDB:Database Password = crimson119"
    adoObject.CommandType = adCmdTable 'Sets The commands to table
    adoObject.RecordSource = adoRecord 'Loads the source from table to records info
    adoObject.Refresh ' refresh the table again
    
End Sub

Public Function IcValid(ByVal stuId As String) As Boolean

    On Error GoTo ErrHandler
    
        Dim tempId As String
        
        frmNewClub.PlayerInfo.Refresh
        
        frmNewClub.PlayerInfo.Recordset.Find ("ID_No = '" & Trim(stuId) & "'")
        tempId = frmNewClub.PlayerInfo.Recordset.Fields("ID_No")
        
        IcValid = True
    Exit Function
    
ErrHandler:
        IcValid = False

End Function

Public Function IcValid2(ByVal stuId As String) As Boolean

    On Error GoTo ErrHandler
    
        Dim tempId As String
        
        AddPlayer.Player.Refresh
        
        AddPlayer.Player.Recordset.Find ("ID_No = '" & Trim(stuId) & "'")
        tempId = AddPlayer.Player.Recordset.Fields("ID_No")
        
        IcValid2 = True
    Exit Function
    
ErrHandler:
        IcValid2 = False

End Function

Public Sub verify_attempt(ByVal log_attem As String)

    Dim end_app As Boolean
    
    If log_attem < 0 Then
        end_app = True
        Unload AdminLogin
    End If
    
End Sub

Public Sub missing()

    MsgBox "Please Enter the Required Field !!", vbCritical + vbOKOnly, "SysMan"
    
End Sub

Public Sub Record_Logins(ByVal adName As String, ByVal logDate As Date)

On Error Resume Next
    
    Call DataConn(AdminLogin.adoLogs, "AdminLog")
    
        AdminLogin.adoLogs.Refresh
        AdminLogin.adoLogs.Recordset.AddNew
    
        With AdminLogin.adoLogs.Recordset
            .Fields(0) = tempUser
            .Fields(1) = logDate
        End With
        
        AdminLogin.adoLogs.Recordset.Update
        AdminLogin.adoLogs.Refresh
        AdminLogin.adoLogs.Recordset.MoveFirst

    
End Sub

Public Sub Record_Logout(ByVal adName As String, ByVal logDate As Date)

On Error Resume Next

    Call DataConn(AdminLogin.adoLogs, "AdminLog")
    
        AdminLogin.adoLogs.Refresh
        'AdminLogin.adoLogs.Recordset.Update
                                
            With AdminLogin.adoLogs.Recordset
                .Fields(2) = logDate
            End With
            
            AdminLogin.adoLogs.Recordset.Update
            AdminLogin.adoLogs.Refresh
            'AdminLogin.adoLogs.Recordset.MoveFirst
 
End Sub

Public Function AppDir() As String
    If Right$(App.Path, 1) = "\" Then
        AppDir = App.Path
    Else
        AppDir = App.Path & "\"
    End If
End Function


