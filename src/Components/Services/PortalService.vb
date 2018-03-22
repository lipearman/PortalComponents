
Public Class PortalServices
    Public Shared connectionString As String = System.Configuration.ConfigurationSettings.AppSettings("PortalConnectionString")
    Public Shared PortalID As String = System.Configuration.ConfigurationSettings.AppSettings("PortalID")

    'Public Shared Function WS() As PortalService.PortalService
    '    Dim TA As New PortalService.PortalService
    '    TA.Url = System.Configuration.ConfigurationSettings.AppSettings("WSPortalService")
    '    TA.Credentials = System.Net.CredentialCache.DefaultCredentials
    '    Return TA
    'End Function

    'Public Shared Sub MessageAlert(ByRef aspxPage As System.Web.UI.Page, _
    '                   ByVal strMessage As String, ByVal strKey As String)

    '    Dim strScript As String = "<script language=JavaScript>alert('" & strMessage & "')</script>"

    '    If (Not aspxPage.IsStartupScriptRegistered(strKey)) Then
    '        aspxPage.RegisterStartupScript(strKey, strScript)
    '    End If
    'End Sub


    'Public Shared Sub WindowOpen(ByRef aspxPage As System.Web.UI.Page, _
    '                           ByVal strURL As String, ByVal strKey As String)

    '    Dim strScript As String = "<script>window.open('" & strURL & "')</script>"

    '    If (Not aspxPage.IsStartupScriptRegistered(strKey)) Then
    '        aspxPage.RegisterStartupScript(strKey, strScript)
    '    End If
    'End Sub


    'Public Shared Function convertSqlDataReaderToDataSet(ByVal reader As System.Data.SqlClient.SqlDataReader) As DataSet
    '    Dim ds As DataSet = New DataSet
    '    Dim schema As DataTable
    '    Dim data As DataTable
    '    Dim i As Integer
    '    Dim dr As DataRow
    '    Dim dc As DataColumn
    '    Dim columnName As String

    '    Do
    '        schema = reader.GetSchemaTable()
    '        data = New DataTable
    '        If Not schema Is Nothing Then
    '            For i = 0 To schema.Rows.Count - 1
    '                dr = schema.Rows(i)
    '                columnName = dr("ColumnName")
    '                If data.Columns.Contains(columnName) Then
    '                    columnName = columnName + "_" + i.ToString()
    '                End If
    '                dc = New DataColumn(columnName, CType(dr("DataType"), Type))
    '                data.Columns.Add(dc)
    '            Next

    '            ds.Tables.Add(data)

    '            While reader.Read()
    '                dr = data.NewRow()
    '                For i = 0 To reader.FieldCount - 1
    '                    dr(i) = reader.GetValue(i)
    '                Next
    '                data.Rows.Add(dr)
    '            End While
    '        Else
    '            dc = New DataColumn("RowsAffected")
    '            data.Columns.Add(dc)
    '            ds.Tables.Add(data)
    '            dr = data.NewRow()
    '            dr(0) = reader.RecordsAffected
    '            data.Rows.Add(dr)
    '        End If
    '    Loop While reader.NextResult()

    '    reader.Close()

    '    Return ds
    'End Function

End Class



