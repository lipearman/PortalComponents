
Partial Public Class DataClasses_PortalDataContextExt
    Inherits DataClasses_PortalDataContext
    Private Shared ReadOnly Property ConnectionString() As String
        Get
            Dim Conn As String = PortalServices.connectionString

            Return Conn
        End Get
    End Property
    Public Sub New()
        MyBase.New(ConnectionString)
    End Sub
End Class