Imports System
Imports System.Configuration
Imports System.Web
Imports System.Data
Imports System.Web.Caching
Imports System.Data.SqlClient
Imports System.Collections


Public Class Configuration

    '
    ' PORTAL
    '

    '*********************************************************************
    '
    ' UpdatePortalInfo() Method <a name="UpdatePortalInfo"></a>
    '
    ' The UpdatePortalInfo method updates the name and access settings for the portal.
    ' These settings are stored in the Xml file PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Sub UpdatePortalInfo(ByVal portalId As Integer, ByVal portalName As String, ByVal alwaysShow As Boolean)
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Get first record of the "Global" element 
        'Dim globalRow As SiteConfiguration._GlobalRow = siteSettings._Global.FindByPortalId(portalId)

        '' Update the values
        'globalRow.PortalId = portalId
        'globalRow.PortalName = portalName
        'globalRow.AlwaysShowEditButton = alwaysShow

        '' Save the changes 
        'SaveSiteSettings()


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _value = dc.PortalCfg_Globals.Where(Function(x) x.PortalId = portalId).FirstOrDefault()
            If Not _value Is Nothing Then
                '_value.PortalId = portalId
                _value.PortalName = portalName
                _value.AlwaysShowEditButton = alwaysShow
                dc.SubmitChanges()
            End If
        End Using


    End Sub

    '
    ' TABS
    '

    '*********************************************************************
    '
    ' AddTab Method <a name="AddTab"></a>
    '
    ' The AddTab method adds a new tab to the portal.  These settings are 
    ' stored in the Xml file PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Function AddTab(ByVal portalId As Integer, ByVal tabName As String, ByVal tabOrder As Integer) As Integer
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Create a new TabRow from the Tab table
        'Dim NewRow As SiteConfiguration.TabRow = siteSettings.Tab.NewTabRow()

        '' Set the properties on the new row
        'NewRow.TabName = tabName
        'NewRow.TabOrder = tabOrder
        'NewRow.ParentId = 0
        'NewRow.AccessRoles = "All Users;"

        '' Add the new TabRow to the Tab table
        'siteSettings.Tab.AddTabRow(NewRow)

        '' Save the changes 
        'SaveSiteSettings()

        '' Return the new TabID
        'Return NewRow.TabId

        Dim _data As New PortalCfg_Tab() With {.PortalId = portalId, .TabName = tabName, .TabOrder = tabOrder, .AccessRoles = "All Users;"}
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            dc.PortalCfg_Tabs.InsertOnSubmit(_data)
            dc.SubmitChanges()
        End Using
        Return _data.TabId

    End Function


    '*********************************************************************
    '
    ' UpdateTab Method <a name="UpdateTab"></a>
    '
    ' The UpdateTab method updates the settings for the specified tab. 
    ' These settings are stored in the Xml file PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Sub UpdateTab(ByVal portalId As Integer, ByVal tabId As Integer, ByVal tabName As String, ByVal tabOrder As Integer, ByVal authorizedRoles As String, ByVal mobileTabName As String, ByVal showMobile As Boolean, ByVal parentId As Integer)

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate tab in the Tab table and set the properties
        'With siteSettings.Tab.FindByTabId(tabId)

        '    .TabName = tabName
        '    .TabOrder = tabOrder
        '    .AccessRoles = authorizedRoles
        '    .MobileTabName = mobileTabName
        '    .ShowMobile = showMobile
        '    .ParentId = parentId

        'End With

        '' Save the changes 
        'SaveSiteSettings()


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _value = dc.PortalCfg_Tabs.Where(Function(x) x.TabId = tabId And x.PortalId = portalId).FirstOrDefault()
            If Not _value Is Nothing Then
                With _value

                    .TabName = tabName
                    .TabOrder = tabOrder
                    .AccessRoles = authorizedRoles
                    .MobileTabName = mobileTabName
                    .ShowMobile = showMobile
                    '.ParentId = parentId

                End With
                dc.SubmitChanges()
            End If
        End Using

    End Sub

    '*********************************************************************
    '
    ' UpdateTabOrder Method <a name="UpdateTabOrder"></a>
    '
    ' The UpdateTabOrder method changes the position of the tab with respect
    ' to other tabs in the portal.  These settings are stored in the Xml 
    ' file PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Sub UpdateTabOrder(ByVal tabId As Integer, ByVal tabOrder As Integer)
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate tab in the Tab table and set the property
        'Dim tabRow As SiteConfiguration.TabRow = siteSettings.Tab.FindByTabId(tabId)

        'tabRow.TabOrder = tabOrder

        '' Save the changes 
        'SaveSiteSettings()

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _value = dc.PortalCfg_Tabs.Where(Function(x) x.TabId = tabId).FirstOrDefault()
            If Not _value Is Nothing Then
                With _value
                    _value.TabOrder = tabOrder
                End With
                dc.SubmitChanges()
            End If
        End Using

    End Sub

    '*********************************************************************
    '
    ' DeleteTab Method <a name="DeleteTab"></a>
    '
    ' The DeleteTab method deletes the selected tab and its modules from 
    ' the settings which are stored in the Xml file PortalCfg.xml.  This 
    ' method also deletes any data from the database associated with all 
    ' modules within this tab.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '	  + <a href="DeleteModule.htm" style="color:green">DeleteModule stored procedure</a>
    '
    '*********************************************************************
    Public Sub DeleteTab(ByVal tabId As Integer)
        '
        ' Delete the Tab in the XML file
        '

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate tab in the Tab table
        'Dim tabTable As SiteConfiguration.TabDataTable = siteSettings.Tab
        'Dim tabRow As SiteConfiguration.TabRow = siteSettings.Tab.FindByTabId(tabId)

        ''
        '' Delete information in the Database relating to each Module being deleted
        ''

        '' Create Instance of Connection and Command Object
        'Dim myConnection As SqlConnection = New SqlConnection(PortalServices.connectionString)
        'Dim myCommand As SqlCommand = New SqlCommand("Portal_DeleteModule", myConnection)

        '' Mark the Command as a SPROC
        'myCommand.CommandType = CommandType.StoredProcedure

        '' Add Parameters to SPROC
        'Dim parameterModuleID As SqlParameter = New SqlParameter("@ModuleID", SqlDbType.Int, 4)
        'myConnection.Open()


        'Dim moduleRow As SiteConfiguration._ModuleRow
        'For Each moduleRow In tabRow.GetModuleRows()
        '    myCommand.Parameters.Clear()
        '    parameterModuleID.Value = moduleRow.ModuleId
        '    myCommand.Parameters.Add(parameterModuleID)

        '    ' Open the database connection and execute the command
        '    myCommand.ExecuteNonQuery()
        'Next

        '' Close the connection
        'myConnection.Close()

        '' Finish removing the Tab row from the Xml file
        'tabTable.RemoveTabRow(tabRow)

        '' Save the changes 
        'SaveSiteSettings()


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _modules = dc.PortalCfg_Modules.Where(Function(x) x.TabId = tabId).ToList()
            For Each x In _modules
                dc.PortalCfg_Modules.DeleteOnSubmit(x)
            Next
            dc.SubmitChanges()

            Dim _tabs = dc.PortalCfg_Tabs.Where(Function(x) x.TabId = tabId).FirstOrDefault()
            If Not _tabs Is Nothing Then
                dc.PortalCfg_Tabs.DeleteOnSubmit(_tabs)
                dc.SubmitChanges()
            End If

        End Using


    End Sub

    '*********************************************************************
    '
    ' UpdateModuleOrder Method  <a name="UpdateModuleOrder"></a>
    '
    ' The UpdateModuleOrder method updates the order in which the modules
    ' in a tab are displayed.  These settings are stored in the Xml file 
    ' PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Sub UpdateModuleOrder(ByVal ModuleId As Integer, ByVal ModuleOrder As Integer, ByVal pane As String)
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate Module in the Module table and update the properties
        'Dim moduleRow As SiteConfiguration._ModuleRow = siteSettings._Module.FindByModuleId(ModuleId)

        'moduleRow.ModuleOrder = ModuleOrder
        'moduleRow.PaneName = pane

        '' Save the changes 
        'SaveSiteSettings()

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _modules = dc.PortalCfg_Modules.Where(Function(x) x.ModuleId = ModuleId).FirstOrDefault()
            If Not _modules Is Nothing Then
                _modules.ModuleOrder = ModuleOrder
                _modules.PaneName = pane
                dc.SubmitChanges()
            End If
        End Using

    End Sub


    '*********************************************************************
    '
    ' AddModule Method  <a name="AddModule"></a>
    '
    ' The AddModule method adds Portal Settings for a new Module within
    ' a Tab.  These settings are stored in the Xml file PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Function AddModule(ByVal tabId As Integer, ByVal moduleOrder As Integer, ByVal paneName As String, ByVal title As String, ByVal moduleDefId As Integer, ByVal cacheTime As Integer, ByVal editRoles As String, ByVal showMobile As Boolean) As Integer

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Create a new ModuleRow from the Module table
        'Dim newModule As SiteConfiguration._ModuleRow = siteSettings._Module.New_ModuleRow()

        '' Set the properties on the new Module
        'With newModule
        '    .ModuleDefId = moduleDefId
        '    .ModuleOrder = moduleOrder
        '    .ModuleTitle = title
        '    .PaneName = paneName
        '    .EditRoles = editRoles
        '    .CacheTimeout = cacheTime
        '    .ShowMobile = showMobile
        '    .TabRow = siteSettings.Tab.FindByTabId(tabId)

        'End With

        '' Add the new ModuleRow to the Module table
        'siteSettings._Module.Add_ModuleRow(newModule)

        '' Save the changes
        'SaveSiteSettings()

        '' Return the new Module ID
        'Return newModule.ModuleId

        Dim newModule As New PortalCfg_Module
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)

            With newModule
                .ModuleDefId = moduleDefId
                .ModuleOrder = moduleOrder
                .ModuleTitle = title
                .PaneName = paneName
                .EditRoles = editRoles
                .CacheTimeout = cacheTime
                .ShowMobile = showMobile
                .TabId = tabId

            End With


            dc.PortalCfg_Modules.InsertOnSubmit(newModule)
            dc.SubmitChanges()
        End Using
        Return newModule.ModuleId



    End Function



    '*********************************************************************
    '
    ' UpdateModule Method  <a name="UpdateModule"></a>
    '
    ' The UpdateModule method updates the Portal Settings for an existing 
    ' Module within a Tab.  These settings are stored in the Xml file
    ' PortalCfg.xml.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Function UpdateModule(ByVal moduleId As Integer, ByVal moduleOrder As Integer, ByVal paneName As String, ByVal title As String, ByVal cacheTime As Integer, ByVal editRoles As String, ByVal showMobile As Boolean) As Integer

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate Module in the Module table and update the properties
        'With siteSettings._Module.FindByModuleId(moduleId)

        '    .ModuleOrder = moduleOrder
        '    .ModuleTitle = title
        '    .PaneName = paneName
        '    .CacheTimeout = cacheTime
        '    .EditRoles = editRoles
        '    .ShowMobile = showMobile

        'End With

        '' Save the changes 
        'SaveSiteSettings()

        '' Return the existing Module ID
        'Return moduleId


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _value = dc.PortalCfg_Modules.Where(Function(x) x.ModuleId = moduleId).FirstOrDefault()
            If Not _value Is Nothing Then
                With _value
                    .ModuleOrder = moduleOrder
                    .ModuleTitle = title
                    .PaneName = paneName
                    .CacheTimeout = cacheTime
                    .EditRoles = editRoles
                    .ShowMobile = showMobile
                End With
                dc.SubmitChanges()
            End If
        End Using

    End Function

    '*********************************************************************
    '
    ' DeleteModule Method  <a name="DeleteModule"></a>
    '
    ' The DeleteModule method deletes a specified Module from the settings
    ' stored in the Xml file PortalCfg.xml.  This method also deletes any 
    ' data from the database associated with this module.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '	  + <a href="DeleteModule.htm" style="color:green">DeleteModule stored procedure</a>
    '
    '*********************************************************************
    Public Sub DeleteModule(ByVal moduleId As Integer)
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        ''
        '' Delete information in the Database relating to Module being deleted
        ''

        '' Create Instance of Connection and Command Object
        'Dim myConnection As SqlConnection = New SqlConnection(PortalServices.connectionString)
        'Dim myCommand As SqlCommand = New SqlCommand("Portal_DeleteModule", myConnection)

        '' Mark the Command as a SPROC
        'myCommand.CommandType = CommandType.StoredProcedure

        '' Add Parameters to SPROC
        'Dim parameterModuleID As SqlParameter = New SqlParameter("@ModuleID", SqlDbType.Int, 4)
        'myConnection.Open()

        'parameterModuleID.Value = moduleId
        'myCommand.Parameters.Add(parameterModuleID)

        '' Open the database connection and execute the command
        'myCommand.ExecuteNonQuery()
        'myConnection.Close()

        '' Finish removing Module
        'siteSettings._Module.Remove_ModuleRow(siteSettings._Module.FindByModuleId(moduleId))

        '' Save the changes 
        'SaveSiteSettings()


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _data = dc.PortalCfg_Modules.Where(Function(x) x.ModuleId = moduleId).FirstOrDefault()
            If Not _data Is Nothing Then
                dc.PortalCfg_Modules.DeleteOnSubmit(_data)
                dc.SubmitChanges()
            End If
        End Using

    End Sub


    '*********************************************************************
    '
    ' UpdateModuleSetting Method  <a name="UpdateModuleSetting"></a>
    '
    ' The UpdateModuleSetting Method updates a single module setting 
    ' in the configuration file.  If the value passed in is String.Empty,
    ' the Setting element is deleted if it exists.  If not, either a 
    ' matching Setting element is updated, or a new Setting element is 
    ' created.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    'Public Sub UpdateModuleSetting(ByVal moduleId As Integer, ByVal key As String, ByVal val As String)
    '    ' Obtain SiteSettings from Current Context
    '    Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

    '    ' Find the appropriate Module in the Module table
    '    Dim moduleRow As SiteConfiguration._ModuleRow = siteSettings._Module.FindByModuleId(moduleId)

    '    ' Find the first (only) settings element
    '    Dim settingsRow As SiteConfiguration.SettingsRow

    '    If moduleRow.GetSettingsRows().Length > 0 Then
    '        settingsRow = moduleRow.GetSettingsRows()(0)
    '    Else
    '        ' Add new settings element
    '        settingsRow = siteSettings.Settings.NewSettingsRow()

    '        ' Set the parent relationship
    '        settingsRow.ModuleRow = moduleRow

    '        siteSettings.Settings.AddSettingsRow(settingsRow)
    '    End If

    '    ' Find the child setting elements
    '    Dim settingRow As SiteConfiguration.SettingRow

    '    Dim settingRows() As SiteConfiguration.SettingRow = settingsRow.GetSettingRows()

    '    If settingRows.Length = 0 Then
    '        ' If there are no Setting elements at all, add one with the new name and value,
    '        ' but only if the value is not empty
    '        If val <> String.Empty Then
    '            settingRow = siteSettings.Setting.NewSettingRow()

    '            ' Set the parent relationship and data
    '            settingRow.SettingsRow = settingsRow
    '            settingRow.Name = key
    '            settingRow.Setting_Text = val

    '            siteSettings.Setting.AddSettingRow(settingRow)
    '        End If
    '    Else
    '        ' Update existing setting element if it matches
    '        Dim found As Boolean = False
    '        Dim i As Int32

    '        ' Find which row matches the input parameter "key" and update the
    '        ' value.  If the value is String.Empty, however, delete the row.
    '        For i = 0 To settingRows.Length - 1 Step i + 1
    '            If settingRows(i).Name = key Then
    '                If val = String.Empty Then
    '                    ' Delete the row
    '                    siteSettings.Setting.RemoveSettingRow(settingRows(i))
    '                Else
    '                    ' Update the value
    '                    settingRows(i).Setting_Text = val
    '                End If

    '                found = True
    '            End If
    '        Next

    '        If found = False Then
    '            ' Setting elements exist, however, there is no matching Setting element.
    '            ' Add one with new name and value, but only if the value is not empty
    '            If val <> String.Empty Then
    '                settingRow = siteSettings.Setting.NewSettingRow()

    '                ' Set the parent relationship and data
    '                settingRow.SettingsRow = settingsRow
    '                settingRow.Name = key
    '                settingRow.Setting_Text = val

    '                siteSettings.Setting.AddSettingRow(settingRow)
    '            End If
    '        End If
    '    End If

    '    ' Save the changes 
    '    SaveSiteSettings()
    'End Sub

    '*********************************************************************
    '
    ' GetModuleSettings Method  <a name="GetModuleSettings"></a>
    '
    ' The GetModuleSettings Method returns a hashtable of custom,
    ' module-specific settings from the configuration file.  This method is
    ' used by some user control modules (Xml, Image, etc) to access misc
    ' settings.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************

    Public Shared Function GetModuleSettings(ByVal moduleId As Integer) As Hashtable
        ' Create a new Hashtable
        Dim _settingsHT As Hashtable = New Hashtable

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate Module in the Module table
        'Dim moduleRow As SiteConfiguration._ModuleRow = siteSettings._Module.FindByModuleId(moduleId)

        '' Find the first (only) settings element
        'If moduleRow.GetSettingsRows().Length > 0 Then
        '    Dim settingsRow As SiteConfiguration.SettingsRow = moduleRow.GetSettingsRows()(0)

        '    If Not settingsRow Is Nothing Then
        '        ' Find the child setting elements and add to the hashtable
        '        Dim sRow As SiteConfiguration.SettingRow
        '        For Each sRow In settingsRow.GetSettingRows()
        '            _settingsHT(sRow.Name) = sRow.Setting_Text
        '        Next
        '    End If
        'End If
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _value = dc.PortalCfg_Modules.Where(Function(x) x.ModuleId = moduleId).FirstOrDefault()
            If Not _value Is Nothing Then
                _settingsHT(_value.ModuleTitle) = _value.ModuleTitle
            End If

        End Using
        Return _settingsHT
    End Function

    Public Function GetModuleDefinitions(ByVal portalId As Integer) As List(Of PortalCfg_ModuleDefinition)

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate Module in the Module table
        'Return siteSettings.ModuleDefinition.Select()

        Dim _value As New List(Of PortalCfg_ModuleDefinition)
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            _value = (From c In dc.PortalCfg_ModuleDefinitions Select c).ToList()
        End Using

        Return _value
    End Function

    '*********************************************************************
    '
    ' AddModuleDefinition() Method <a name="AddModuleDefinition"></a>
    '
    ' The AddModuleDefinition add the definition for a new module type
    ' to the portal.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************

    Public Function AddModuleDefinition(ByVal portalId As Integer, ByVal name As String, ByVal desktopSrc As String, ByVal mobileSrc As String) As Integer

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Create new ModuleDefinitionRow
        'Dim newModuleDef As SiteConfiguration.ModuleDefinitionRow = siteSettings.ModuleDefinition.NewModuleDefinitionRow()

        '' Set the parameter values
        'With newModuleDef

        '    .FriendlyName = name
        '    .DesktopSourceFile = desktopSrc
        '    .MobileSourceFile = mobileSrc

        'End With

        '' Add the new ModuleDefinitionRow to the ModuleDefinition table
        'siteSettings.ModuleDefinition.AddModuleDefinitionRow(newModuleDef)

        '' Save the changes
        'SaveSiteSettings()

        '' Return the new ModuleDefID
        'Return newModuleDef.ModuleDefId


        Dim _data As New PortalCfg_ModuleDefinition
        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)

            With _data

                .ModuleDefName = name
                .ModuleDefSourceFile = desktopSrc
                .ModuleDefMobileSourceFile = mobileSrc

            End With

            dc.PortalCfg_ModuleDefinitions.InsertOnSubmit(_data)
            dc.SubmitChanges()
        End Using
        Return _data.ModuleDefId



    End Function

    '*********************************************************************
    '
    ' DeleteModuleDefinition() Method <a name="DeleteModuleDefinition"></a>
    '
    ' The DeleteModuleDefinition method deletes the specified module type 
    ' definition from the portal.  Each module which is related to the
    ' ModuleDefinition is deleted from each tab in the configuration
    ' file, and all data relating to each module is deleted from the
    ' database.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '    + <a href="DeleteModule.htm" style="color:green">DeleteModule Stored Procedure</a>
    '
    '*********************************************************************
    Public Sub DeleteModuleDefinition(ByVal defId As Integer)
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        ''
        '' Delete information in the Database relating to each Module being deleted
        ''

        '' Create Instance of Connection and Command Object
        'Dim myConnection As SqlConnection = New SqlConnection(PortalServices.connectionString)
        'Dim myCommand As SqlCommand = New SqlCommand("Portal_DeleteModule", myConnection)

        '' Mark the Command as a SPROC
        'myCommand.CommandType = CommandType.StoredProcedure

        '' Add Parameters to SPROC
        'Dim parameterModuleID As SqlParameter = New SqlParameter("@ModuleID", SqlDbType.Int, 4)
        'myConnection.Open()

        'Dim moduleRow As SiteConfiguration._ModuleRow
        'For Each moduleRow In siteSettings._Module.Select()
        '    If moduleRow.ModuleDefId = defId Then
        '        myCommand.Parameters.Clear()
        '        parameterModuleID.Value = moduleRow.ModuleId
        '        myCommand.Parameters.Add(parameterModuleID)

        '        ' Delete the xml module associated with the ModuleDef
        '        ' in the configuration file
        '        siteSettings._Module.Remove_ModuleRow(moduleRow)

        '        ' Open the database connection and execute the command
        '        myCommand.ExecuteNonQuery()
        '    End If
        'Next

        'myConnection.Close()

        '' Finish removing Module Definition
        'siteSettings.ModuleDefinition.RemoveModuleDefinitionRow(siteSettings.ModuleDefinition.FindByModuleDefId(defId))

        '' Save the changes 
        'SaveSiteSettings()


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            Dim _modules = dc.PortalCfg_Modules.Where(Function(x) x.ModuleDefId = defId).ToList()
            For Each x In _modules
                dc.PortalCfg_Modules.DeleteOnSubmit(x)
            Next
            dc.SubmitChanges()

            Dim _ModuleDefinitions = dc.PortalCfg_ModuleDefinitions.Where(Function(x) x.ModuleDefId = defId).FirstOrDefault()
            If Not _ModuleDefinitions Is Nothing Then
                dc.PortalCfg_ModuleDefinitions.DeleteOnSubmit(_ModuleDefinitions)
                dc.SubmitChanges()
            End If

        End Using


    End Sub

    '*********************************************************************
    '
    ' UpdateModuleDefinition() Method <a name="UpdateModuleDefinition"></a>
    '
    ' The UpdateModuleDefinition method updates the settings for the 
    ' specified module type definition.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Sub UpdateModuleDefinition(ByVal defId As Integer, ByVal name As String, ByVal desktopSrc As String, ByVal mobileSrc As String)

        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate Module in the Module table and update the properties
        'With siteSettings.ModuleDefinition.FindByModuleDefId(defId)

        '    .FriendlyName = name
        '    .DesktopSourceFile = desktopSrc
        '    .MobileSourceFile = mobileSrc

        'End With

        '' Save the changes 
        'SaveSiteSettings()


        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)     
            Dim _ModuleDefinitions = dc.PortalCfg_ModuleDefinitions.Where(Function(x) x.ModuleDefId = defId).FirstOrDefault()
            If Not _ModuleDefinitions Is Nothing Then
                With _ModuleDefinitions

                    .ModuleDefName = name
                    .ModuleDefSourceFile = desktopSrc
                    .ModuleDefMobileSourceFile = mobileSrc

                End With
                dc.SubmitChanges()
            End If

        End Using

    End Sub

    '*********************************************************************
    '
    ' GetSingleModuleDefinition Method
    '
    ' The GetSingleModuleDefinition method returns a ModuleDefinitionRow
    ' object containing details about a specific module definition in the
    ' configuration file.
    '
    ' Other relevant sources:
    '    + <a href="#SaveSiteSettings" style="color:green">SaveSiteSettings() method</a>
    '	  + <a href="PortalCfg.xml" style="color:green">PortalCfg.xml</a>
    '
    '*********************************************************************
    Public Function GetSingleModuleDefinition(ByVal defId As Integer) As PortalCfg_ModuleDefinition
        '' Obtain SiteSettings from Current Context
        'Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Items("SiteSettings"), SiteConfiguration)

        '' Find the appropriate Module in the Module table
        'Return siteSettings.ModuleDefinition.FindByModuleDefId(defId)

        Dim _ModuleDefinitions As New PortalCfg_ModuleDefinition

        Using dc As New DataClasses_PortalDataContext(PortalServices.connectionString)
            _ModuleDefinitions = dc.PortalCfg_ModuleDefinitions.Where(Function(x) x.ModuleDefId = defId).FirstOrDefault()
        End Using
        Return _ModuleDefinitions
    End Function

    '*********************************************************************
    '
    ' GetSiteSettings Static Method
    '
    ' The Configuration.GetSiteSettings Method returns a typed
    ' dataset of the all of the site configuration settings from the
    ' XML configuration file.  This method is used in Global.asax to
    ' push the settings into the current HttpContext, so that all of the 
    ' pages, content modules and classes throughout the rest of the request
    ' may access them.
    '
    ' The SiteConfiguration object is cached using the ASP.NET Cache API,
    ' with a file-change dependency on the XML configuration file.  Normallly,
    ' this method just returns a copy of the object in the cache.  When the
    ' configuration is updated and changes are saved to the the XML file,
    ' the SiteConfiguration object is evicted from the cache.  The next time 
    ' this method runs, it will read from the XML file again and insert a
    ' fresh copy of the SiteConfiguration into the cache.
    '
    '*********************************************************************
    'Public Shared Function GetSiteSettings() As SiteConfiguration
    '    Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Cache("SiteSettings"), SiteConfiguration)

    '    ' If the SiteConfiguration isn't cached, load it from the XML file and add it into the cache.
    '    If siteSettings Is Nothing Then
    '        ' Create the dataset
    '        siteSettings = New SiteConfiguration

    '        ' Retrieve the location of the XML configuration file
    '        Dim configFile As String = HttpContext.Current.Server.MapPath(ConfigurationSettings.AppSettings("configFile"))

    '        ' Set the AutoIncrement property to true for easier adding of rows
    '        siteSettings.Tab.TabIdColumn.AutoIncrement = True
    '        siteSettings._Module.ModuleIdColumn.AutoIncrement = True
    '        siteSettings.ModuleDefinition.ModuleDefIdColumn.AutoIncrement = True

    '        ' Load the XML data into the DataSet
    '        siteSettings.ReadXml(configFile)

    '        ' Store the dataset in the cache
    '        HttpContext.Current.Cache.Insert("SiteSettings", siteSettings, New CacheDependency(configFile))
    '    End If

    '    Return siteSettings
    'End Function

    '*********************************************************************
    '
    ' SaveSiteSettings Method <a name="SaveSiteSettings"></a>
    '
    ' The Configuration.SaveSiteSettings overwrites the the XML file with the
    ' settings in the SiteConfiguration object in context.  The object will in 
    ' turn be evicted from the cache and be reloaded from the XML file the next
    ' time GetSiteSettings() is called.
    '
    '*********************************************************************
    'Public Sub SaveSiteSettings()
    '    ' Obtain SiteSettings from the Cache
    '    Dim siteSettings As SiteConfiguration = CType(HttpContext.Current.Cache("SiteSettings"), SiteConfiguration)
    '    ' Check the object
    '    If siteSettings Is Nothing Then
    '        ' If SaveSiteSettings() is called once, the cache is cleared.  If it is
    '        ' then called again before Global.Application_BeginRequest is called, 
    '        ' which reloads the cache, the siteSettings object will be Null 
    '        siteSettings = GetSiteSettings()
    '    End If
    '    Dim configFile As String = HttpContext.Current.Server.MapPath(ConfigurationSettings.AppSettings("configFile"))
    '    ' Object is evicted from the Cache here.  
    '    siteSettings.WriteXml(configFile)
    'End Sub
End Class


