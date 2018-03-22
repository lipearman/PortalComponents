Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Imports System.Text
Imports System.Configuration
Imports System.Security


Public Enum EnumProjectRole
    Administrator
    Manager
    Developer
    Tester
End Enum

Public Class ProjectMember
    Private mUserName As String
    Private mDateAssigned As DateTime
    Private mProjectRole As EnumProjectRole

    Public Property UserName() As String
        Get
            Return Me.mUserName
        End Get
        Set(ByVal Value As String)
            Me.mUserName = Value
        End Set
    End Property

    Public Property DateAssigned() As DateTime
        Get
            Return Me.mDateAssigned
        End Get
        Set(ByVal Value As DateTime)
            Me.mDateAssigned = Value
        End Set
    End Property

    Public Property ProjectRole() As EnumProjectRole
        Get
            Return Me.mProjectRole
        End Get
        Set(ByVal Value As EnumProjectRole)
            Me.mProjectRole = Value
        End Set
    End Property

    Public Shared Function GetProjectRoles() As Hashtable
        Dim ht As Hashtable = New Hashtable
        'Dim dr As SqlDataReader = Nothing
        'Try
        '    dr = SqlHelper.ExecuteReader(PortalServices.connectionString, "GetProjectRoles", PortalSecurity.UserName)
        '    While dr.Read()
        '        ht.Add(dr("ProjectID"), dr("ProjectRole"))
        '    End While
        'Finally
        '    dr.Close()
        'End Try

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _user = (From c In dc.Portal_Users Where c.UserName = PortalSecurity.UserName).FirstOrDefault()
            If Not _user Is Nothing Then
                Dim _userrole = (From c In dc.Portal_UserRoles Where c.UserID = _user.UserID).ToList()
                For Each _item In _userrole
                    Dim RoleID = _item.RoleID
                    Dim _role = dc.Portal_Roles.Where(Function(x) x.RoleID = RoleID).FirstOrDefault()
                    ht.Add(_role.RoleID, _role.RoleCode)
                Next

            End If
        End Using

        Return ht
    End Function


    Public Overloads Overrides Function Equals(ByVal o As Object) As Boolean
        If (Not o Is Nothing) And (TypeOf (o) Is ProjectMember) Then
            Dim p As ProjectMember = CType(o, ProjectMember)
            Return p.UserName = Me.UserName
        End If
        Return False
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return Me.UserName.GetHashCode()
    End Function




End Class
