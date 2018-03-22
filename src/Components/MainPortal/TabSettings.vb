

'*********************************************************************
'
' TabSettings Class
'
' Class that encapsulates the detailed settings for a specific Tab 
' in the Portal
'
'*********************************************************************
Public Class TabSettings
    Public TabIndex As Integer
    Public TabId As Integer
    Public TabName As String
    Public TabOrder As Integer
    Public AuthorizedRoles As ArrayList = New ArrayList
    Public Modules As ArrayList = New ArrayList
    Public MobileTabName As String
    Public ShowMobile As Boolean
    Public ParentId As Integer?
    Public PageID As String
    Public IconID As String
End Class

