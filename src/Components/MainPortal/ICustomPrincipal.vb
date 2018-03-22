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

Public Interface ICustomPrincipal
    Inherits IPrincipal
    Function IsInProjectRole(ByVal projectID As Integer, ByVal ParamArray roleArray As EnumProjectRole()) As Boolean
    Function GetProjectRole(ByVal projectID As Integer) As EnumProjectRole
End Interface


