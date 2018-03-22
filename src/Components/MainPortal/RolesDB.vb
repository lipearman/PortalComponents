Imports System
Imports System.Configuration
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections

'*********************************************************************
'
' RolesDB Class
'
' Class that encapsulates all data logic necessary to add/query/delete
' Users, Roles and security settings values within the Portal database.
'
'*********************************************************************
Public Class RolesDB
    '*********************************************************************
    '
    ' GetPortalRoles() Method <a name="GetPortalRoles"></a>
    '
    ' The GetPortalRoles method returns a list of all role names for the 
    ' specified portal.
    '
    ' Other relevant sources:
    '     + <a href="GetRolesByUser.htm" style="color:green">GetPortalRoles Stored Procedure</a>
    '
    '*********************************************************************

    Public Function GetPortalRoles(ByVal portalId As Integer) As List(Of Portal_Role)
        Dim _value As New List(Of Portal_Role)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            _value = dc.Portal_Roles.Where(Function(x) x.PortalID = portalId).ToList()
        End Using
        Return _value
    End Function

    '*********************************************************************
    '
    ' AddRole() Method <a name="AddRole"></a>
    '
    ' The AddRole method creates a new security role for the specified portal,
    ' and returns the new RoleID value.
    '
    ' Other relevant sources:
    '     + <a href="AddRole.htm" style="color:green">AddRole Stored Procedure</a>
    '
    '*********************************************************************

    Public Function AddRole(ByVal portalId As Integer, ByVal roleName As String) As Integer
        Dim _data As New Portal_Role() With {.PortalID = portalId, .RoleName = roleName}
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            dc.Portal_Roles.InsertOnSubmit(_data)
            dc.SubmitChanges()
        End Using
        Return _data.RoleID
    End Function

    '*********************************************************************
    '
    ' DeleteRole() Method <a name="DeleteRole"></a>
    '
    ' The DeleteRole deletes the specified role from the portal database.
    '
    ' Other relevant sources:
    '     + <a href="DeleteRole.htm" style="color:green">DeleteRole Stored Procedure</a>
    '
    '*********************************************************************

    Public Sub DeleteRole(ByVal roleId As Integer)
       Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _data = dc.Portal_Roles.Where(Function(x) x.RoleID = roleId).FirstOrDefault()
            If Not _data Is Nothing Then
                dc.Portal_Roles.DeleteOnSubmit(_data)
                dc.SubmitChanges()
            End If
        End Using
    End Sub

    '*********************************************************************
    '
    ' UpdateRole() Method <a name="UpdateRole"></a>
    '
    ' The UpdateRole method updates the friendly name of the specified role.
    '
    ' Other relevant sources:
    '     + <a href="UpdateRole.htm" style="color:green">UpdateRole Stored Procedure</a>
    '
    '*********************************************************************

    Public Sub UpdateRole(ByVal roleId As Integer, ByVal roleName As String)
       Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _data = dc.Portal_Roles.Where(Function(x) x.RoleID = roleId).FirstOrDefault()
            If Not _data Is Nothing Then
                _data.RoleName = roleName
                dc.SubmitChanges()
            End If
        End Using
    End Sub


    '
    ' USER ROLES
    '

    '*********************************************************************
    '
    ' GetRoleMembers() Method <a name="GetRoleMembers"></a>
    '
    ' The GetRoleMembers method returns a list of all members in the specified
    ' security role.
    '
    ' Other relevant sources:
    '     + <a href="GetRoleMembers.htm" style="color:green">GetRoleMembers Stored Procedure</a>
    '
    '*********************************************************************

    Public Function GetRoleMembers(ByVal roleId As Integer) As List(Of Portal_User)
        Dim _value As New List(Of Portal_User)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)         
            Dim _roleuser = dc.Portal_UserRoles.Where(Function(x) x.RoleID = roleId).ToList()
            If _roleuser.Count > 0 Then
                Dim _userroles = _roleuser.Select(Function(x) x.UserID).ToList()
                _value = dc.Portal_Users.Where(Function(x) _userroles.Contains(x.UserID)).ToList()
            End If
        End Using
        Return _value
    End Function

    '*********************************************************************
    '
    ' AddUserRole() Method <a name="AddUserRole"></a>
    '
    ' The AddUserRole method adds the user to the specified security role.
    '
    ' Other relevant sources:
    '     + <a href="AddUserRole.htm" style="color:green">AddUserRole Stored Procedure</a>
    '
    '*********************************************************************
    Public Sub AddUserRole(ByVal roleId As Integer, ByVal userId As Integer)        
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            dc.Portal_UserRoles.InsertOnSubmit(New Portal_UserRole() With {.RoleID = roleId, .UserID = userId})
            dc.SubmitChanges()
        End Using
    End Sub

    '*********************************************************************
    '
    ' DeleteUserRole() Method <a name="DeleteUserRole"></a>
    '
    ' The DeleteUserRole method deletes the user from the specified role.
    '
    ' Other relevant sources:
    '     + <a href="DeleteUserRole.htm" style="color:green">DeleteUserRole Stored Procedure</a>
    '
    '*********************************************************************

    Public Sub DeleteUserRole(ByVal roleId As Integer, ByVal userId As Integer)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _data = dc.Portal_UserRoles.Where(Function(x) x.RoleID = roleId And x.UserID = userId).FirstOrDefault()
            If Not _data Is Nothing Then
                dc.Portal_UserRoles.DeleteOnSubmit(_data)
                dc.SubmitChanges()
            End If
        End Using
    End Sub
    '
    ' USERS
    '

    '*********************************************************************
    '
    ' GetUsers() Method <a name="GetUsers"></a>
    '
    ' The GetUsers method returns returns the UserID, Name and Email for 
    ' all registered users.
    '
    ' Other relevant sources:
    '     + <a href="GetUsers.htm" style="color:green">GetUsers Stored Procedure</a>
    '
    '*********************************************************************

    Public Function GetUsers() As List(Of Portal_User)
        Dim _value As New List(Of Portal_User)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            _value = (From c In dc.Portal_Users Select c).ToList()
        End Using
        Return _value
    End Function

End Class




