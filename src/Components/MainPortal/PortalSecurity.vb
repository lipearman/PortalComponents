Imports System
Imports System.Collections
Imports System.Configuration
Imports System.Web
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security.Principal


'*********************************************************************
'
' PortalSecurity Class
'
' The PortalSecurity class encapsulates two helper methods that enable
' developers to easily check the role status of the current browser client.
'
'*********************************************************************

Public Class PortalSecurity


    Public Shared ReadOnly Property UserName() As String
        Get
            Return HttpContext.Current.User.Identity.Name
        End Get
    End Property
    '*********************************************************************
    '
    ' Security.Encrypt() Method
    '
    ' The Encrypt method encrypts a clean string into a hashed string
    '
    '*********************************************************************

    Public Shared Function Encrypt(ByVal cleanString As String) As String
        Dim clearBytes() As Byte = New UnicodeEncoding().GetBytes(cleanString)
        Dim hashedBytes() As Byte = (CType(CryptoConfig.CreateFromName("MD5"), HashAlgorithm)).ComputeHash(clearBytes)

        Return BitConverter.ToString(hashedBytes)

    End Function

    '*********************************************************************
    '
    ' PortalSecurity.HasEditPermissions() Method
    '
    ' The HasEditPermissions method enables developers to easily check 
    ' whether the current browser client has access to edit the settings
    ' of a specified portal module
    '
    '*********************************************************************

    Public Shared Function HasEditPermissions(ByVal moduleId As Integer) As Boolean
        Dim accessRoles As New ArrayList
        Dim tabid As Integer = 0
        'Dim editRoles As ArrayList

        ' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)
        Try
            ' Find the appropriate Module in the Module table
            'Dim moduleRow As SiteConfiguration._ModuleRow = siteSettings._Module.FindByModuleId(moduleId)
            'accessRoles = ""
            'editRoles = ""

            Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
                Dim _Module = dc.PortalCfg_Modules.Where(Function(x) x.ModuleId = moduleId).FirstOrDefault()
                If Not _Module Is Nothing Then
                    tabid = _Module.TabId

                    'editRoles = _Module.EditRoles
                    Dim _Roles = dc.Portal_TabRoles.Where(Function(x) x.TabID = _Module.TabId).ToArray()
                    'If Not _tab Is Nothing Then
                    '    accessRoles = _tab.AccessRoles
                    'End If
                    For Each _role In _Roles
                        accessRoles.Add(_role)
                    Next
                End If
            End Using




            If PortalSecurity.IsInRoles(accessRoles, tabid) = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function


    '*********************************************************************
    '
    ' PortalSecurity.IsInRole() Method
    '
    ' The IsInRole method enables developers to easily check the role
    ' status of the current browser client.
    '
    '*********************************************************************
    Public Shared Function IsInRole(ByVal role As String) As Boolean

        Return HttpContext.Current.User.IsInRole(role)
    End Function


    '*********************************************************************
    '
    ' PortalSecurity.IsInRoles() Method
    '
    ' The IsInRoles method enables developers to easily check the role
    ' status of the current browser client against an array of roles
    '
    '*********************************************************************
    ' Public Shared Function IsInRoles(ByVal roles As ArrayList) As Boolean
    Public Shared Function IsInRoles(ByVal roles As ArrayList, ByVal TabId As Integer) As Boolean
        If roles Is Nothing Then
            Return True
        Else
            'Dim context As HttpContext = HttpContext.Current
            'Dim role As String
            'For Each role In roles.Split(New Char() {";"c})
            '    'If role <> "" And role <> Nothing And CheckRole(role) Then 'And ((role = "All Users") Or (context.User.IsInRole(role))) Then
            '    '    Return True
            '    'End If

            '    If CheckRole(role) Then
            '        Return True
            '    End If
            'Next

            Dim userObj As UsersDB = New UsersDB()
            Dim userRolesObj = userObj.GetRolesByUserTabs(TabId, HttpContext.Current.User.Identity.Name)


            For Each RoleID In userRolesObj
                If roles.Contains(RoleID) Then
                    Return True
                End If
            Next
        End If



        Return False
    End Function


    Private Shared Function CheckRole(ByVal _CheckRole) As Boolean
        Dim userObj As UsersDB = New UsersDB()
        Dim userRolesObj = userObj.GetRolesByUser(HttpContext.Current.User.Identity.Name)
        For Each x In userRolesObj
            If (_CheckRole.LastIndexOf(x.RoleCode)) > -1 Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Function IsInProjectRole(ByVal projectID As Integer, ByVal ParamArray roleArray As EnumProjectRole()) As Boolean
        Dim custom As ICustomPrincipal = CType(HttpContext.Current.User, ICustomPrincipal)
        Return custom.IsInProjectRole(projectID, roleArray)
    End Function

    Public Shared Function GetProjectRole(ByVal projectID As Integer) As EnumProjectRole
        Dim custom As ICustomPrincipal = CType(HttpContext.Current.User, ICustomPrincipal)
        Return custom.GetProjectRole(projectID)
    End Function


End Class



