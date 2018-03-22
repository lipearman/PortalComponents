Imports System
Imports System.IO
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls


'*********************************************************************
'
' PortalModuleControl Class
'
' The PortalModuleControl class defines a custom base class inherited by all
' desktop portal modules within the Portal.
' 
' The PortalModuleControl class defines portal specific properties
' that are used by the portal framework to correctly display portal modules
'
'*********************************************************************

Public Class PortalModuleControl
    Inherits UserControl

    ' Private field variables
    Public PortalConn As String = ConfigurationSettings.AppSettings("PortalConnectionString")

    Private _moduleConfiguration As Portal.Components.ModuleSettings
    Private _isEditable As Integer = 0
    Private _portalId As Integer = 0
    Private _settings As Hashtable

    ' Public property accessors

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property ModuleId() As Integer
        Get
            Return CType(_moduleConfiguration.ModuleId, Integer)
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Property PortalId() As Integer
        Get
            Return _portalId
        End Get
        Set(ByVal Value As Integer)
            _portalId = Value
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property IsEditable() As Boolean
        Get

            ' Perform tri-state switch check to avoid having to perform a security
            ' role lookup on every property access (instead caching the result)

            If _isEditable = 0 Then

                ' Obtain PortalSettings from Current Context

                Dim portalSettings As PortalSettings = CType(HttpContext.Current.Items("PortalSettings"), PortalSettings)

                If portalSettings.AlwaysShowEditButton = True Or PortalSecurity.IsInRoles(_moduleConfiguration.AuthorizedEditRoles, _moduleConfiguration.TabId) Then
                    _isEditable = 1

                Else
                    _isEditable = 2
                End If
            End If

            Return (_isEditable = 1)
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Property ModuleConfiguration() As Portal.Components.ModuleSettings
        Get
            Return _moduleConfiguration
        End Get
        Set(ByVal Value As Portal.Components.ModuleSettings)
            _moduleConfiguration = Value
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Settings() As Hashtable
        Get

            If _settings Is Nothing Then

                _settings = Configuration.GetModuleSettings(ModuleId)
            End If

            Return _settings
        End Get
    End Property
End Class


