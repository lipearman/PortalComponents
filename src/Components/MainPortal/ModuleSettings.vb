
'*********************************************************************
'
' ModuleSettings Class
'
' Class that encapsulates the detailed settings for a specific Tab 
' in the Portal.  ModuleSettings implements the IComparable interface 
' so that an ArrayList of ModuleSettings objects may be sorted by
' ModuleOrder, using the ArrayList's Sort() method.
'
'*********************************************************************

Public Class ModuleSettings
    Implements IComparable

    Public ModuleId As Integer
    Public ModuleDefId As Integer
    Public TabId As Integer
    Public CacheTimeOut As Integer
    Public ModuleOrder As Integer
    Public PaneName As String
    Public ModuleTitle As String
    Public AuthorizedEditRoles As ArrayList

    Public DesktopSourceFile As String
    Public ShowMobile As Boolean
    Public MobileSrc As String


    Public Function CompareTo(ByVal value As Object) As Integer Implements IComparable.CompareTo

        If value Is Nothing Then
            Return 1
        End If

        Dim compareOrder As Integer = (CType(value, ModuleSettings)).ModuleOrder

        If Me.ModuleOrder = compareOrder Then
            Return 0
        End If
        If Me.ModuleOrder < compareOrder Then
            Return -1
        End If
        If Me.ModuleOrder > compareOrder Then
            Return 1
        End If
        Return 0
    End Function
End Class


