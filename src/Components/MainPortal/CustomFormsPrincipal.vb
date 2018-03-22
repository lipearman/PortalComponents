Imports System
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security.Principal

Public Class CustomFormsPrincipal
    Inherits GenericPrincipal
    Implements ICustomPrincipal
    Private projectRoles As Hashtable

    Public Sub New(ByVal identity As IIdentity, ByVal roles() As String, ByVal projectRoles As Hashtable)
        MyBase.New(identity, roles)
        Me.projectRoles = projectRoles
    End Sub
    Public Function IsInProjectRole(ByVal projectID As Integer, ByVal ParamArray roleArray As EnumProjectRole()) As Boolean Implements ICustomPrincipal.IsInProjectRole
        Dim projectRole As EnumProjectRole
        For Each projectRole In roleArray
            If projectRoles(projectID).ToString() = projectRole.ToString() Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function GetProjectRole(ByVal projectID As Integer) As EnumProjectRole Implements ICustomPrincipal.GetProjectRole
        Dim p As EnumProjectRole = CType(projectRoles(projectID).ToString(), EnumProjectRole)
        Return p
    End Function
End Class


