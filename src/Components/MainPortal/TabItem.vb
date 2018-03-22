
'*********************************************************************
'
' TabItem Class
'
' This class encapsulates the basic attributes of a Tab, and is used
' by the administration pages when manipulating tabs.  TabItem implements 
' the IComparable interface so that an ArrayList of TabItems may be sorted
' by TabOrder, using the ArrayList's Sort() method.
'
'*********************************************************************

Public Class TabItem
    Implements IComparable

    Private _tabOrder As Integer
    Private _name As String
    Private _id As Integer

    Public Property TabOrder() As Integer
        Get
            Return _tabOrder
        End Get
        Set(ByVal Value As Integer)
            _tabOrder = Value
        End Set
    End Property

    Public Property TabName() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Value
        End Set
    End Property

    Public Property TabId() As Integer
        Get
            Return _id
        End Get
        Set(ByVal Value As Integer)
            _id = Value
        End Set
    End Property

    Public Function CompareTo(ByVal value As Object) As Integer Implements IComparable.CompareTo

        If value Is Nothing Then
            Return 1
        End If

        Dim compareOrder As Integer = (CType(value, TabItem)).TabOrder

        If Me.TabOrder = compareOrder Then
            Return 0
        End If
        If Me.TabOrder < compareOrder Then
            Return -1
        End If
        If Me.TabOrder > compareOrder Then
            Return 1
        End If
        Return 0
    End Function
End Class

