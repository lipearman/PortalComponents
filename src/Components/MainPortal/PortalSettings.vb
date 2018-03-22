
Public Class PortalSettings

    Public PortalId As Integer?
    Public SortPath As Integer
    Public PortalName As String
    Public AlwaysShowEditButton As Boolean
    Public DesktopTabs As New List(Of DesktopTab)
    Public MobileTabs As New ArrayList
    Public ActiveTab As New TabSettings
    Public PageID As String
    Public IconID As String
    '*********************************************************************
    '
    ' PortalSettings Constructor
    '
    ' The PortalSettings Constructor encapsulates all of the logic
    ' necessary to obtain configuration settings necessary to render
    ' a Portal Tab view for a given request.
    '
    ' These Portal Settings are stored within PortalCFG.xml, and are
    ' fetched below by calling config.GetSiteSettings().
    ' The method config.GetSiteSettings() fills the SiteConfiguration
    ' class, derived from a DataSet, which PortalSettings accesses.
    '       
    '*********************************************************************
    'Public Sub New(ByVal tabIndex As Integer, ByVal tabId As Integer)
    Public Sub New(ByVal _PageID As String)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)

            Me.PortalId = PortalServices.PortalID
            Dim _Portal_cfg = (From c In dc.PortalCfg_Globals Where c.PortalId.Equals(Me.PortalId)).FirstOrDefault()
            If _Portal_cfg IsNot Nothing Then
                Me.PortalName = _Portal_cfg.PortalName
            End If



            Me.PageID = _PageID

            ' Get the configuration data
            Dim config As Configuration = New Configuration
            'Dim siteSettings As SiteConfiguration = config.GetSiteSettings()
            Dim _tabs = (From c In dc.v_DesktopTabs Where c.PortalId = Me.PortalId Or c.TabId.Equals(1) Select c Order By c.Sortpath).ToList()


            Dim tabDetails As New DesktopTab

            'With tabDetails
            '    .TabId = 1
            '    .TabName = "Home"
            '    .TabOrder = 1
            '    '.AccessRoles = ""
            '    .Sortpath = "0001"
            'End With

            'Me.DesktopTabs.Add(tabDetails)

            ' Read the Desktop Tab Information, and sort by Tab Order

            For Each item In _tabs
                tabDetails = New DesktopTab

                With tabDetails

                    .TabId = item.TabId
                    .TabName = item.TabName
                    .TabOrder = item.TabOrder

                    .ParentId = item.ParentId
                    .PortalId = item.PortalId
                    .Sortpath = item.Sortpath

                    .PageID = item.PageId.ToString()
                    If Not String.IsNullOrEmpty(item.IconID) Then
                        .IconID = item.IconID.ToString()
                    End If



                    Dim _TabsRoles = (From c In dc.Portal_TabRoles Where c.TabID.Equals(item.TabId) Select c.RoleID).ToList
                    For Each irole In _TabsRoles
                        .AuthorizedRoles.Add(irole)
                    Next




                End With

                Me.DesktopTabs.Add(tabDetails)
            Next

            ' If the PortalSettings.ActiveTab property is set to 0, change it to  
            ' the TabID of the first tab in the DesktopTabs collection
            'If Me.ActiveTab.TabId = 0 Then
            '    Me.ActiveTab.TabId = CType(Me.DesktopTabs(0), DesktopTab).PageId
            'End If


            '' Read the Mobile Tab Information, and sort by Tab Order
            'Dim mRow As SiteConfiguration.TabRow
            'For Each mRow In siteSettings.Tab.Select("ShowMobile='true'", "TabOrder")
            '    Dim tabDetails As New PortalCfg_Tab

            '    With tabDetails
            '        .TabId = mRow.TabId
            '        .TabName = mRow.MobileTabName
            '        .AccessRoles = mRow.AccessRoles
            '    End With

            '    Me.MobileTabs.Add(tabDetails)
            'Next

            ' Read the Module Information for the current (Active) tab

            Dim _Page = (From c In dc.PortalCfg_Tabs Where c.PageID.Equals(Me.PageID)).FirstOrDefault()

            If _Page Is Nothing Then Return

            Dim _modules = (From c In dc.v_ModuleSettings Where c.TabId.Equals(_Page.TabId) Order By c.ModuleOrder).ToList()

            ' Get Modules for this Tab based on the Data Relation
            For Each moduleRow In _modules
                Dim moduleSettings As New Portal.Components.ModuleSettings

                With moduleSettings

                    .ModuleTitle = moduleRow.ModuleTitle
                    .ModuleId = moduleRow.ModuleId
                    .ModuleDefId = moduleRow.ModuleDefId
                    .ModuleOrder = moduleRow.ModuleOrder
                    .TabId = _Page.TabId
                    .PaneName = moduleRow.PaneName
                    '.EditRoles = moduleRow.EditRoles
                    .CacheTimeout = moduleRow.CacheTimeout
                    '.ModuleDefCode = moduleRow.ModuleDefCode
                    .DesktopSourceFile = moduleRow.DesktopSourceFile


                    'Try
                    '    .ShowMobile = moduleRow.ShowMobile
                    'Catch ex As Exception
                    '    .ShowMobile = False
                    'End Try


                    ' ModuleDefinition data
                    'Dim modDefRow As SiteConfiguration.ModuleDefinitionRow = siteSettings.ModuleDefinition.FindByModuleDefId(.ModuleDefId)

                    '.DesktopSrc = modDefRow.DesktopSourceFile
                    'Try
                    '    .MobileSrc = modDefRow.MobileSourceFile
                    'Catch ex As Exception
                    '    .MobileSrc = ""
                    'End Try



                    'Dim _data = dc.PortalCfg_ModuleDefinitions.Where(Function(x) x.ModuleDefId = .ModuleDefId).FirstOrDefault()
                    'If Not _data Is Nothing Then
                    '.DesktopSourceFile = moduleRow.DesktopSourceFile
                    'Try
                    '    .MobileSourceFile = moduleRow.MobileSourceFile
                    'Catch ex As Exception
                    '    .MobileSourceFile = ""
                    'End Try
                    'End If



                End With

                Me.ActiveTab.Modules.Add(moduleSettings)
            Next

            ' Sort the modules in order of ModuleOrder
            'Me.ActiveTab.Modules.Sort()

            ' Get the first row in the Global table
            Dim globalSettings = (From c In dc.PortalCfg_Globals Where c.PortalId = Me.PortalId).FirstOrDefault()

            ' Read Portal global settings 
            Me.PortalId = globalSettings.PortalId
            Me.PortalName = globalSettings.PortalName




            Me.AlwaysShowEditButton = IIf(globalSettings.AlwaysShowEditButton Is Nothing, False, True)
            Me.ActiveTab.TabIndex = _Page.TabOrder
            Me.ActiveTab.TabId = _Page.TabId

            Dim activeTab = (From c In dc.PortalCfg_Tabs Where c.TabId.Equals(_Page.TabId)).FirstOrDefault()


            If Not activeTab Is Nothing Then

                Me.ActiveTab.TabOrder = activeTab.TabOrder
                Me.ActiveTab.ParentId = activeTab.ParentId
                'Try
                '    Me.ActiveTab.MobileTabName = activeTab.MobileTabName
                'Catch ex As Exception
                '    Me.ActiveTab.MobileTabName = activeTab.TabName
                'End Try

                'Me.ActiveTab.AuthorizedRoles = activeTab.AccessRoles
                Dim _ActiveRoles = (From c In dc.Portal_TabRoles Where c.TabID.Equals(_Page.TabId) Select c.RoleID).ToList()
                For Each irole In _ActiveRoles
                    Me.ActiveTab.AuthorizedRoles.Add(irole)
                Next

                Me.ActiveTab.TabName = activeTab.TabName
                'Try
                '    Me.ActiveTab.ShowMobile = activeTab.ShowMobile
                'Catch ex As Exception
                '    Me.ActiveTab.ShowMobile = False
                'End Try
                Me.PageID = activeTab.PageID
                Me.ActiveTab.PageID = activeTab.PageID
                Me.IconID = activeTab.IconID
                Me.ActiveTab.IconID = activeTab.IconID

            End If

        End Using
    End Sub

End Class


