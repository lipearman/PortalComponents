Imports System.Data
Imports System.Data.SqlClient

'*********************************************************************
'
' UsersDB Class
'
' The UsersDB class encapsulates all data logic necessary to add/login/query
' users within the Portal Users database.
'
' Important Note: The UsersDB class is only used when forms-based cookie
' authentication is enabled within the portal.  When windows based
' authentication is used instead, then either the Windows SAM or Active Directory
' is used to store and validate all username/password credentials.
'
'*********************************************************************
Public Class UsersDB
    '*********************************************************************
    '
    ' UsersDB.AddUser() Method <a name="AddUser"></a>
    '
    ' The AddUser method inserts a new user record into the "Users" database table.
    '
    ' Other relevant sources:
    '     + <a href="AddUser.htm" style="color:green">AddUser Stored Procedure</a>
    '
    '*********************************************************************
    Public Function AddUser(ByVal _data As Portal_User) As Integer
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            dc.Portal_Users.InsertOnSubmit(_data)
            dc.SubmitChanges()
        End Using
        Return _data.UserID
    End Function
    '*********************************************************************
    '
    ' UsersDB.DeleteUser() Method <a name="DeleteUser"></a>
    '
    ' The DeleteUser method deleted a  user record from the "Users" database table.
    '
    ' Other relevant sources:
    '     + <a href="DeleteUser.htm" style="color:green">DeleteUser Stored Procedure</a>
    '
    '*********************************************************************
    Public Sub DeleteUser(ByVal userId As Integer)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _data = dc.Portal_Users.Where(Function(x) x.UserID = userId).FirstOrDefault()
            If Not _data Is Nothing Then
                dc.Portal_Users.DeleteOnSubmit(_data)
                dc.SubmitChanges()
            End If
        End Using
    End Sub
    '*********************************************************************
    '
    ' UsersDB.UpdateUser() Method <a name="DeleteUser"></a>
    '
    ' The UpdateUser method deleted a  user record from the "Users" database table.
    '
    ' Other relevant sources:
    '     + <a href="UpdateUser.htm" style="color:green">UpdateUser Stored Procedure</a>
    '
    '*********************************************************************
    Public Sub UpdateUser(ByVal _data As Portal_User)
        If _data Is Nothing Then Return

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _value = dc.Portal_Users.Where(Function(x) x.UserID = _data.UserID).FirstOrDefault()
            If Not _value Is Nothing Then
                _value.Email = _data.Email
                _value.UserName = _data.UserName
                _value.Password = _data.Password
                dc.SubmitChanges()
            End If
        End Using
    End Sub
    '*********************************************************************
    '
    ' UsersDB.GetRolesByUser() Method <a name="GetRolesByUser"></a>
    '
    ' The DeleteUser method deleted a  user record from the "Users" database table.
    '
    ' Other relevant sources:
    '     + <a href="GetRolesByUser.htm" style="color:green">GetRolesByUser Stored Procedure</a>
    '
    '*********************************************************************
    Public Function GetRolesByUser(ByVal UserName As String) As List(Of Portal_Role)
        Dim _value As New List(Of Portal_Role)

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _user = dc.Portal_Users.Where(Function(x) x.UserName = UserName).FirstOrDefault()
            If Not _user Is Nothing Then
                Dim _roleuser = dc.Portal_UserRoles.Where(Function(x) x.UserID = _user.UserID).ToList()
                If _roleuser.Count > 0 Then
                    Dim _userroles = _roleuser.Select(Function(x) x.RoleID).ToList()
                    _value = dc.Portal_Roles.Where(Function(x) _userroles.Contains(x.RoleID)).ToList()
                End If
            End If
        End Using
        Return _value
    End Function

    Public Function GetRolesByUserTabs(ByVal TabId As Integer, ByVal UserName As String) As List(Of Integer)
        Dim _value As New List(Of Integer)

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _user = dc.Portal_Users.Where(Function(x) x.UserName = UserName).FirstOrDefault()

            If _user IsNot Nothing Then
                Dim _userTabs = dc.Portal_UserTabs.Where(Function(x) x.USERID = _user.UserID And x.TABID.Equals(TabId)).FirstOrDefault()
                Dim _userRoles = dc.Portal_UserRoles.Where(Function(x) x.UserID = _user.UserID).ToList()

                If _userTabs IsNot Nothing And _userRoles.Count > 0 Then
                    Dim _Roles = (From c In _userRoles Select c.RoleID).ToList() 
                    Dim _tabroles = dc.Portal_TabRoles.Where(Function(x) x.TabID = _userTabs.TABID And _Roles.Contains(x.RoleID)).ToList()

                    If _tabroles.Count > 0 Then
                        _value = _tabroles.Select(Function(x) x.RoleID).ToList()
                    End If
                End If
            End If

        End Using
        Return _value
    End Function
    '*********************************************************************
    '
    ' GetSingleUser Method
    '
    ' The GetSingleUser method returns a SqlDataReader containing details
    ' about a specific user from the Users database table.
    '
    '*********************************************************************
    Public Function GetSingleUser(ByVal UserName As String) As Portal_User
        Dim _user As New Portal_User
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            _user = dc.Portal_Users.Where(Function(x) x.UserName = UserName).FirstOrDefault()
        End Using
        Return _user
    End Function
    '*********************************************************************
    '
    ' GetRoles() Method <a name="GetRoles"></a>
    '
    ' The GetRoles method returns a list of role names for the user.
    '
    ' Other relevant sources:
    '     + <a href="GetRolesByUser.htm" style="color:green">GetRolesByUser Stored Procedure</a>
    '
    '*********************************************************************
    Public Function GetRoles(ByVal UserName As String) As List(Of Portal_Role)
        Dim _value As New List(Of Portal_Role)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _user = dc.Portal_Users.Where(Function(x) x.UserName = UserName).FirstOrDefault()
            If Not _user Is Nothing Then
                Dim _roleuser = dc.Portal_UserRoles.Where(Function(x) x.UserID = _user.UserID).ToList()
                If _roleuser.Count > 0 Then
                    Dim _userroles = _roleuser.Select(Function(x) x.RoleID).ToList()
                    _value = dc.Portal_Roles.Where(Function(x) _userroles.Contains(x.RoleID)).ToList()
                End If
            End If
        End Using
        Return _value
    End Function
    '*********************************************************************
    '
    ' UsersDB.Login() Method <a name="Login"></a>
    '
    ' The Login method validates a email/password pair against credentials
    ' stored in the users database.  If the email/password pair is valid,
    ' the method returns user's name.
    '
    ' Other relevant sources:
    '     + <a href="UserLogin.htm" style="color:green">UserLogin Stored Procedure</a>
    '
    '*********************************************************************
    Public Function Login(ByVal UserName As String, ByVal password As String) As Portal_User
        Dim _value As New Portal_User
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            _value = dc.Portal_Users.Where(Function(x) x.UserName = UserName And x.Password = password).FirstOrDefault()
        End Using
        Return _value
    End Function

End Class


