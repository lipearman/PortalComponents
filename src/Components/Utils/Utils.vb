Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Security.Cryptography
Imports System.DirectoryServices

Public Class Utils
    Structure objCookie
        Dim Username As String
        Dim Dep As String
        Dim Name As String
        Dim Level As String
        Dim Broker As String
        Dim Agent As String
        Dim IsBroker As String
        Dim IsPrintPDF As String
    End Structure


    Public Shared Sub ConvertDataSetToExcel(ByVal ds As DataSet, ByVal response As System.Web.HttpResponse, ByVal FileName As String)
        'first let's clean up the response.object
        response.Clear()
        response.AddHeader("content-disposition", "attachment;filename=" & FileName)
        response.Charset = String.Empty
        response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)
        response.ContentEncoding = System.Text.Encoding.Default
        'set the response mime type for excel
        response.ContentType = "application/vnd.ms-excel"
        'create a string writer
        Dim stringWrite As New System.IO.StringWriter
        'create an htmltextwriter which uses the stringwriter
        Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)
        'instantiate a datagrid
        Dim dg As New System.Web.UI.WebControls.DataGrid
        'set the datagrid datasource to the dataset passed in
        dg.DataSource = ds.Tables(0)

        'bind the datagrid
        dg.DataBind()
        'tell the datagrid to render itself to our htmltextwriter
        dg.RenderControl(htmlWrite)
        'all that's left is to output the html
        response.Flush()
        response.Write(stringWrite.ToString)
        response.End()


    End Sub
    Public Shared Sub ConvertDataSetToExcel(ByVal ds As DataSet, ByVal TableIndex As Integer, ByVal response As System.Web.HttpResponse, ByVal FileName As String)
        'lets make sure a table actually exists at the passed in value
        'if it is not call the base method
        If TableIndex > ds.Tables.Count - 1 Then
            ConvertDataSetToExcel(ds, response, FileName)
        End If
        'we've got a good table so
        'let's clean up the response.object
        response.Clear()
        response.AddHeader("content-disposition", "attachment;filename=" & FileName)
        response.Charset = String.Empty
        response.ContentEncoding = System.Text.Encoding.Default

        response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)

        'set the response mime type for excel
        response.ContentType = "application/vnd.ms-excel"
        'create a string writer
        Dim stringWrite As New System.IO.StringWriter
        'create an htmltextwriter which uses the stringwriter
        Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)
        'instantiate a datagrid
        Dim dg As New System.Web.UI.WebControls.DataGrid
        'set the datagrid datasource to the dataset passed in
        dg.DataSource = ds.Tables(TableIndex)
        'bind the datagrid
        dg.DataBind()
        'tell the datagrid to render itself to our htmltextwriter
        dg.RenderControl(htmlWrite)
        'all that's left is to output the html
        response.Flush()
        response.Write(stringWrite.ToString)
        response.End()


    End Sub
    Public Shared Sub ConvertDataSetToExcel(ByVal ds As DataSet, ByVal TableName As String, ByVal response As System.Web.HttpResponse, ByVal FileName As String)
        'let's make sure the table name exists
        'if it does not then call the default method
        If ds.Tables(TableName) Is Nothing Then
            ConvertDataSetToExcel(ds, response, FileName)
        End If
        'we've got a good table so
        'let's clean up the response.object
        response.Clear()
        response.AddHeader("content-disposition", "attachment;filename=" & FileName)
        response.Charset = String.Empty
        response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)

        response.ContentEncoding = System.Text.Encoding.Default
        'set the response mime type for excel
        response.ContentType = "application/vnd.ms-excel"
        'create a string writer
        Dim stringWrite As New System.IO.StringWriter
        'create an htmltextwriter which uses the stringwriter
        Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)
        'instantiate a datagrid
        Dim dg As New System.Web.UI.WebControls.DataGrid
        'set the datagrid datasource to the dataset passed in
        dg.DataSource = ds.Tables(TableName)
        'bind the datagrid
        dg.DataBind()
        'tell the datagrid to render itself to our htmltextwriter
        dg.RenderControl(htmlWrite)
        'all that's left is to output the html
        response.Flush()
        response.Write(stringWrite.ToString)
        response.End()


    End Sub
    Public Shared Sub MessageAlert(ByRef aspxPage As System.Web.UI.Page, ByVal strMessage As String, ByVal strKey As String)

        Dim strScript As String = "<script language=JavaScript>alert('" & strMessage & "')</script>"

        If (Not aspxPage.IsStartupScriptRegistered(strKey)) Then
            aspxPage.RegisterStartupScript(strKey, strScript)
        End If
    End Sub
    Public Shared Sub WindowOpen(ByRef aspxPage As System.Web.UI.Page, ByVal strURL As String, ByVal strKey As String)

        Dim strScript As String = "<script>window.open('" & strURL & "')</script>"

        If (Not aspxPage.IsStartupScriptRegistered(strKey)) Then
            aspxPage.RegisterStartupScript(strKey, strScript)
        End If
    End Sub
    Public Shared Sub WindowModalOpen(ByRef aspxPage As System.Web.UI.Page, ByVal strURL As String, ByVal strKey As String)

        Dim strScript As String = "<script>window.showModalDialog('" & strURL & "',null,'status:no;dialogWidth:800px;dialogHeight:600px;dialogHide:true;help:no;scroll:yes');</script>"

        If (Not aspxPage.IsStartupScriptRegistered(strKey)) Then
            aspxPage.RegisterStartupScript(strKey, strScript)
        End If
    End Sub
    Public Shared Sub WindowModalOpen(ByRef aspxPage As System.Web.UI.Page, ByVal strURL As String, ByVal strKey As String, ByVal pxWidth As Integer, ByVal pxHeight As Integer)

        Dim strScript As String = "<script>window.showModalDialog('" & strURL & "',null,'status:no;dialogWidth:" & pxWidth & "px;dialogHeight:" & pxHeight & "px;dialogHide:true;help:no;scroll:yes');</script>"

        If (Not aspxPage.IsStartupScriptRegistered(strKey)) Then
            aspxPage.RegisterStartupScript(strKey, strScript)
        End If
    End Sub
    Public Shared Function CalBirthDate(ByVal Birthdate As Date, ByVal EffectiveDate As Date) As Integer
        Dim GetAdge As Integer = DateDiff(DateInterval.Year, Birthdate, EffectiveDate)

        If (Birthdate.Month > EffectiveDate.Month) Or _
        (Birthdate.Month = EffectiveDate.Month And Birthdate.Day > EffectiveDate.Day) Then
            GetAdge = GetAdge - 1
        End If

        If GetAdge <= 0 Then
            GetAdge = 1
        End If

        Return GetAdge
    End Function
    Public Shared Function RoundUp(ByVal GetNumber As Decimal) As Integer
        If CInt(Mid(Format(GetNumber, "0000000000.0000"), 14, 1)) >= 5 Then
            GetNumber += 0.01
        End If
        If CInt(Mid(Format(GetNumber, "0000000000.0000"), 13, 1)) >= 5 Then
            GetNumber += 0.1
        End If
        Return GetNumber
    End Function
    Public Shared Function MyCDate(ByVal GetDate As String) As Date
        Dim vDate() As String = Split(GetDate, "/", -1, 1)
        'Try
        Return CType(vDate(2) & "-" & vDate(1) & "-" & vDate(0), Date)
        'Catch ex As Exception
        '    Return Nothing
        'End Try
    End Function
    Public Shared Function SQLREADER2DT(ByVal dr As System.Data.SqlClient.SqlDataReader) As DataTable
        'Create New Datatable 
        Dim dt As New DataTable
        'Retrieve Columns and Column Headers
        For i As Integer = 0 To dr.FieldCount - 1
            Dim dcol As New DataColumn
            dcol.ColumnName = dr.GetName(i)
            'Add Each Column in the loop
            dt.Columns.Add(dcol)
        Next
        Do While dr.Read()
            'Create Datarow
            Dim drow As DataRow
            'set row equal to a new row in the datatable
            drow = dt.NewRow()
            For i As Integer = 0 To dr.FieldCount - 1
                'Add All Data From Fields into new row
                drow.Item(i) = dr(i)
            Next
            'add each the new row to the datatable
            dt.Rows.Add(drow)
        Loop
        Return dt
    End Function
    Public Shared Function ConvertDataReaderToDataSet(ByVal reader As System.Data.SqlClient.SqlDataReader) As DataSet
        Dim dataSet As DataSet = New DataSet
        Dim dataRow As DataRow
        Dim columnName As String
        Dim column As DataColumn
        Dim schemaTable As DataTable
        Dim dataTable As DataTable

        Do
            ' Create new data table
            schemaTable = reader.GetSchemaTable
            dataTable = New DataTable
            If Not IsDBNull(schemaTable) Then
                ' A query returning records was executed
                Dim i As Integer
                For i = 0 To schemaTable.Rows.Count - 1
                    dataRow = schemaTable.Rows(i)
                    ' Create a column name that is unique in the data table
                    columnName = dataRow("ColumnName")
                    'Add the column definition to the data table
                    column = New DataColumn(columnName, CType(dataRow("DataType"), Type))
                    dataTable.Columns.Add(column)
                Next
                dataSet.Tables.Add(dataTable)

                'Fill the data table we just created
                While reader.Read()
                    dataRow = dataTable.NewRow()
                    For i = 0 To reader.FieldCount - 1
                        dataRow(i) = reader(i)
                    Next
                    dataTable.Rows.Add(dataRow)
                End While
            Else
                'No records were returned
                column = New DataColumn("RowsAffected")
                dataTable.Columns.Add(column)
                dataSet.Tables.Add(dataTable)
                dataRow = dataTable.NewRow()
                dataRow(0) = reader.RecordsAffected
                dataTable.Rows.Add(dataRow)
            End If
        Loop While reader.NextResult()
        Return dataSet
    End Function
    Public Shared Function MyCheckEmail(ByVal GetEmail As String) As String
        If Trim(GetEmail) = "" Then
            Return "Please fields Email!!!"
            Exit Function
        Else
            Dim strTmp As String, n As Long, sEXT As String
            Dim EMsg As String = "" 'reset on open for good form 
            sEXT = GetEmail
            Do While InStr(1, sEXT, ".") <> 0
                sEXT = Right(sEXT, Len(sEXT) - InStr(1, sEXT, "."))
            Loop
            If GetEmail = "" Then
                Return "You did not enter an email address!"
                Exit Function
            ElseIf InStr(1, GetEmail, "@") = 0 Then
                Return "Your email address does not contain an @ sign."
                Exit Function
            ElseIf InStr(1, GetEmail, "@") = 1 Then
                Return "Your @ sign can not be the first character in your email address!"
                Exit Function
            ElseIf InStr(1, GetEmail, "@") = Len(GetEmail) Then
                Return "Your @sign can not be the last character in your email address!"
                Exit Function
            ElseIf EXTisOK(sEXT) = False Then
                EMsg = EMsg & "Your email address is not carrying a valid ending!"
                EMsg = EMsg & "\nIt must be one of the following..."
                EMsg = EMsg & "\n.com, .net, .gov, .org, .edu, .biz, .tv Or included country\'s assigned country code"
                Return EMsg
                Exit Function
            ElseIf Len(GetEmail) < 6 Then
                Return "Your email address is shorter than 6 characters which is impossible."
                Exit Function
            End If
            strTmp = GetEmail
            Do While InStr(1, strTmp, "@") <> 0
                n = 1
                strTmp = Right(strTmp, Len(strTmp) - InStr(1, strTmp, "@"))
            Loop
            If n > 1 Then
                Return "You have more than 1 @ sign in your email address"
                Exit Function
            End If
        End If

    End Function
    Private Shared Function EXTisOK(ByVal sEXT As String) As Boolean
        Dim EXT As String, X As Long
        EXTisOK = False
        If Left(sEXT, 1) <> "." Then sEXT = "." & sEXT
        sEXT = UCase(sEXT) 'just to avoid errors 
        EXT = EXT & ".COM.EDU.GOV.NET.BIZ.ORG.TV"
        EXT = EXT & ".AF.AL.DZ.As.AD.AO.AI.AQ.AG.AP.AR.AM.AW.AU.AT.AZ.BS.BH.BD.BB.BY"
        EXT = EXT & ".BE.BZ.BJ.BM.BT.BO.BA.BW.BV.BR.IO.BN.BG.BF.MM.BI.KH.CM.CA.CV.KY"
        EXT = EXT & ".CF.TD.CL.CN.CX.CC.CO.KM.CG.CD.CK.CR.CI.HR.CU.CY.CZ.DK.DJ.DM.DO"
        EXT = EXT & ".TP.EC.EG.SV.GQ.ER.EE.ET.FK.FO.FJ.FI.CS.SU.FR.FX.GF.PF.TF.GA.GM.GE.DE"
        EXT = EXT & ".GH.GI.GB.GR.GL.GD.GP.GU.GT.GN.GW.GY.HT.HM.HN.HK.HU.IS.IN.ID.IR.IQ"
        EXT = EXT & ".IE.IL.IT.JM.JP.JO.KZ.KE.KI.KW.KG.LA.LV.LB.LS.LR.LY.LI.LT.LU.MO.MK.MG"
        EXT = EXT & ".MW.MY.MV.ML.MT.MH.MQ.MR.MU.YT.MX.FM.MD.MC.MN.MS.MA.MZ.NA"
        EXT = EXT & ".NR.NP.NL.AN.NT.NC.NZ.NI.NE.NG.NU.NF.KP.MP.NO.OM.PK.PW.PA.PG.PY"
        EXT = EXT & ".PE.PH.PN.PL.PT.PR.QA.RE.RO.RU.RW.GS.SH.KN.LC.PM.ST.VC.SM.SA.SN.SC"
        EXT = EXT & ".SL.SG.SK.SI.SB.SO.ZA.KR.ES.LK.SD.SR.SJ.SZ.SE.CH.SY.TJ.TW.TZ.TH.TG.TK"
        EXT = EXT & ".TO.TT.TN.TR.TM.TC.TV.UG.UA.AE.UK.US.UY.UM.UZ.VU.VA.VE.VN.VG.VI"
        EXT = EXT & ".WF.WS.EH.YE.YU.ZR.ZM.ZW"
        EXT = UCase(EXT) 'just to avoid errors 
        If InStr(1, EXT, sEXT, vbBinaryCompare) <> 0 Then EXTisOK = True
    End Function
    Public Shared Sub Return_Caller(ByRef aspxPage As System.Web.UI.Page)
        Dim strjscript As String
        'oLiteral.Text = ""
        strjscript = "<script language=""javascript"">"
        strjscript = strjscript & "var theFrom = window.opener;"
        strjscript = strjscript & "theFrom.document.Form1.submit();"
        strjscript = strjscript & "this.close();"
        strjscript = strjscript & "</script>"

        'oLiteral.Text = strjscript  'Set the literal control's text to the JScript code
        If (Not aspxPage.IsStartupScriptRegistered("Return_Caller")) Then
            aspxPage.RegisterStartupScript("Return_Caller", strjscript)
        End If

    End Sub

    Public Shared Function objExport2Excel(ByVal DS As DataSet, ByVal TableIndex As Integer, ByRef aspxPage As System.Web.UI.Page, ByVal FileFormatPath As String, ByVal FileName As String) As String
        ''================================
        'Dim excelApplication As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
        'Dim excelWorkbook As Microsoft.Office.Interop.Excel.Workbook ' = CType(excelApplication.Workbooks.Add(System.Reflection.Missing.Value), Microsoft.Office.Interop.Excel.Workbook)
        'Dim excelSheet(DS.Tables.Count - 1) As Microsoft.Office.Interop.Excel.Worksheet '= CType(excelWorkbook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)

        'If FileFormatPath = "" Then
        '    excelWorkbook = CType(excelApplication.Workbooks.Add(System.Reflection.Missing.Value), Microsoft.Office.Interop.Excel.Workbook)
        'Else
        '    excelWorkbook = CType(excelApplication.Workbooks.Open(FileFormatPath), Microsoft.Office.Interop.Excel.Workbook)
        'End If

        'If TableIndex = 0 Then
        '    For ix As Integer = 0 To DS.Tables.Count - 1
        '        excelSheet(ix) = CType(excelWorkbook.Sheets(ix + 1), Microsoft.Office.Interop.Excel.Worksheet)
        '        excelSheet(ix).Name = DS.Tables(ix).TableName
        '        'add the columns
        '        For i As Integer = 0 To DS.Tables(ix).Columns.Count - 1
        '            CType(excelSheet(ix).Cells(1, i + 1), Microsoft.Office.Interop.Excel.Range).Value2 = DS.Tables(ix).Columns(i).ColumnName
        '        Next i

        '        'set the column styles
        '        excelSheet(ix).Range(excelSheet(ix).Cells(1, 1), excelSheet(ix).Cells(1, DS.Tables(ix).Columns.Count)).Font.Bold = True
        '        excelSheet(ix).Range(excelSheet(ix).Cells(1, 1), excelSheet(ix).Cells(1, DS.Tables(ix).Columns.Count)).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        '        For i As Integer = 0 To DS.Tables(ix).Rows.Count - 1
        '            excelSheet(ix).Cells.Range(excelSheet(ix).Cells(i + 2, 1), excelSheet(ix).Cells(i + 2, DS.Tables(ix).Columns.Count)).NumberFormat = "@"
        '            excelSheet(ix).Cells.Range(excelSheet(ix).Cells(i + 2, 1), excelSheet(ix).Cells(i + 2, DS.Tables(ix).Columns.Count)).Value2 = DS.Tables(ix).Rows(i).ItemArray
        '        Next i
        '    Next

        'Else
        '    excelSheet(0) = CType(excelWorkbook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        '    excelSheet(0).Name = DS.Tables(0).TableName
        '    'add the columns
        '    For i As Integer = 0 To DS.Tables(0).Columns.Count - 1
        '        CType(excelSheet(0).Cells(1, i + 1), Microsoft.Office.Interop.Excel.Range).Value2 = DS.Tables(0).Columns(i).ColumnName
        '    Next i

        '    'set the column styles
        '    excelSheet(0).Range(excelSheet(0).Cells(1, 1), excelSheet(0).Cells(1, DS.Tables(0).Columns.Count)).Font.Bold = True
        '    excelSheet(0).Range(excelSheet(0).Cells(1, 1), excelSheet(0).Cells(1, DS.Tables(0).Columns.Count)).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        '    For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
        '        excelSheet(0).Cells.Range(excelSheet(0).Cells(i + 2, 1), excelSheet(0).Cells(i + 2, DS.Tables(0).Columns.Count)).NumberFormat = "@"
        '        excelSheet(0).Cells.Range(excelSheet(0).Cells(i + 2, 1), excelSheet(0).Cells(i + 2, DS.Tables(0).Columns.Count)).Value2 = DS.Tables(0).Rows(i).ItemArray
        '    Next i
        'End If



        'excelApplication.Visible = False
        'excelApplication.DisplayAlerts = False
        'excelWorkbook.Activate()
        ''Dim FileName As String = "DS" & Format(Now, "yyyyMMddHHmmss") & ".xls"
        'FileName = FileName & Format(Now, "yyyyMMddHHmmss") & ".xls"
        'Dim FilePath As String = aspxPage.Server.MapPath("~/DocCRM") & "/" & FileName

        'If Not System.IO.Directory.Exists(aspxPage.Server.MapPath("~/DocCRM")) Then
        '    System.IO.Directory.CreateDirectory(aspxPage.Server.MapPath("~/DocCRM"))
        'End If

        'excelWorkbook.SaveAs(FilePath)
        'excelWorkbook.Close()
        'excelWorkbook = Nothing
        'excelApplication.Quit()
        'excelApplication = Nothing
        'Dim proc As System.Diagnostics.Process
        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next

        'Dim strHdr As String = ""
        'strHdr = "attachment;filename=" & FileName

        'With aspxPage.Response
        '    .Clear()
        '    .ContentType = "application/vnd.ms-excel"
        '    .ContentEncoding = System.Text.Encoding.Default
        '    .AppendHeader("Content-Disposition", strHdr)
        '    .WriteFile(FilePath)
        '    .Flush()
        '    .Clear()
        '    .Close()
        'End With
        'If System.IO.File.Exists(FilePath) Then
        '    System.IO.File.Delete(FilePath)
        'End If

        '''===============================

        '''we've got a good table so
        '''let's clean up the response.object
        ''aspxPage.Response.Clear()
        ''aspxPage.Response.AddHeader("content-disposition", "attachment;filename=" & FileName)
        ''aspxPage.Response.Charset = String.Empty
        ''aspxPage.Response.ContentEncoding = System.Text.Encoding.Default

        ''aspxPage.Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)

        '''set the response mime type for excel
        ''aspxPage.Response.ContentType = "application/vnd.ms-excel"

        '''all that's left is to output the html
        ''aspxPage.Response.Flush()
        ''aspxPage.Response.Write(FilePath)
        ''aspxPage.Response.End()

        ''If System.IO.File.Exists(FilePath) Then
        ''    System.IO.File.Delete(FilePath)
        ''End If
        ''==============================
        ''Dim fs As New System.io.FileStream(FilePath, IO.FileMode.Open, IO.FileAccess.Read)
        ''Dim bw As New System.IO.BinaryReader(fs)
        ''Dim byt() As Byte
        ''byt = bw.ReadBytes(CInt(fs.Length))
        ''aspxPage.Response.ContentType = "application/vnd.ms-excel"
        ''aspxPage.Response.OutputStream.Write(byt, 0, fs.Length)

        ''===================================
        ''aspxPage.Response.Redirect("~/DocCRM/" & FileName)
        Return FileName
    End Function
    Public Shared Sub objExport2Excel(ByVal DS As DataSet, ByVal TableIndex As Integer, ByRef aspxPage As System.Web.UI.Page)
        ''================================
        'Dim excelApplication As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
        'Dim excelWorkbook As Microsoft.Office.Interop.Excel.Workbook = CType(excelApplication.Workbooks.Add(System.Reflection.Missing.Value), Microsoft.Office.Interop.Excel.Workbook)
        'Dim excelSheet(DS.Tables.Count - 1) As Microsoft.Office.Interop.Excel.Worksheet '= CType(excelWorkbook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)


        'If TableIndex = 0 Then
        '    For ix As Integer = 0 To DS.Tables.Count - 1
        '        excelSheet(ix) = CType(excelWorkbook.Sheets(ix + 1), Microsoft.Office.Interop.Excel.Worksheet)
        '        excelSheet(ix).Name = DS.Tables(ix).TableName
        '        'add the columns
        '        For i As Integer = 0 To DS.Tables(ix).Columns.Count - 1
        '            CType(excelSheet(ix).Cells(1, i + 1), Microsoft.Office.Interop.Excel.Range).Value2 = DS.Tables(ix).Columns(i).ColumnName
        '        Next i

        '        'set the column styles
        '        excelSheet(ix).Range(excelSheet(ix).Cells(1, 1), excelSheet(ix).Cells(1, DS.Tables(ix).Columns.Count)).Font.Bold = True
        '        excelSheet(ix).Range(excelSheet(ix).Cells(1, 1), excelSheet(ix).Cells(1, DS.Tables(ix).Columns.Count)).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        '        For i As Integer = 0 To DS.Tables(ix).Rows.Count - 1
        '            excelSheet(ix).Cells.Range(excelSheet(ix).Cells(i + 2, 1), excelSheet(ix).Cells(i + 2, DS.Tables(ix).Columns.Count)).NumberFormat = "@"
        '            excelSheet(ix).Cells.Range(excelSheet(ix).Cells(i + 2, 1), excelSheet(ix).Cells(i + 2, DS.Tables(ix).Columns.Count)).Value2 = DS.Tables(ix).Rows(i).ItemArray
        '        Next i
        '    Next

        'Else
        '    excelSheet(0) = CType(excelWorkbook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        '    excelSheet(0).Name = DS.Tables(0).TableName
        '    'add the columns
        '    For i As Integer = 0 To DS.Tables(0).Columns.Count - 1
        '        CType(excelSheet(0).Cells(1, i + 1), Microsoft.Office.Interop.Excel.Range).Value2 = DS.Tables(0).Columns(i).ColumnName
        '    Next i

        '    'set the column styles
        '    excelSheet(0).Range(excelSheet(0).Cells(1, 1), excelSheet(0).Cells(1, DS.Tables(0).Columns.Count)).Font.Bold = True
        '    excelSheet(0).Range(excelSheet(0).Cells(1, 1), excelSheet(0).Cells(1, DS.Tables(0).Columns.Count)).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        '    For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
        '        excelSheet(0).Cells.Range(excelSheet(0).Cells(i + 2, 1), excelSheet(0).Cells(i + 2, DS.Tables(0).Columns.Count)).NumberFormat = "@"
        '        excelSheet(0).Cells.Range(excelSheet(0).Cells(i + 2, 1), excelSheet(0).Cells(i + 2, DS.Tables(0).Columns.Count)).Value2 = DS.Tables(0).Rows(i).ItemArray
        '    Next i
        'End If



        'excelApplication.Visible = False
        'excelApplication.DisplayAlerts = False
        'excelWorkbook.Activate()

        'Dim FilePath As String = aspxPage.Server.MapPath("~/DocCRM") & "/DS" & Format(Now, "yyyyMMddHHmmss") & ".xls"

        'If Not System.IO.Directory.Exists(aspxPage.Server.MapPath("~/DocCRM")) Then
        '    System.IO.Directory.CreateDirectory(aspxPage.Server.MapPath("~/DocCRM"))
        'End If



        'excelWorkbook.SaveAs(FilePath)
        'excelWorkbook.Close()
        'excelWorkbook = Nothing
        'excelApplication.Quit()
        'excelApplication = Nothing
        'Dim proc As System.Diagnostics.Process
        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next


        'Dim strHdr As String = ""
        'strHdr = "attachment;filename=ExcelFile.xls"

        'With aspxPage.Response
        '    .Clear()
        '    .ContentType = "application/vnd.ms-excel"
        '    .ContentEncoding = System.Text.Encoding.Default
        '    .AppendHeader("Content-Disposition", strHdr)
        '    .WriteFile(FilePath)
        '    .Flush()
        '    .Clear()
        '    .Close()
        'End With
        'If System.IO.File.Exists(FilePath) Then
        '    System.IO.File.Delete(FilePath)
        'End If


        ''Dim fs As New FileStream(FilePath, FileMode.Open,
        ''FileAccess.Read)
        ''Dim bw As New System.IO.BinaryReader(fs)
        ''Dim byt() As Byte
        ''byt = bw.ReadBytes(CInt(fs.Length))
        ''Response.ContentType = "application/vnd.ms-excel"
        ''Response.OutputStream.Write(byt, 0, length)

    End Sub
    Public Shared Sub objExport2Excel(ByVal DS As DataSet, ByVal TableName As String, ByRef aspxPage As System.Web.UI.Page)
        '''================================
        'Dim excelApplication As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
        'Dim excelWorkbook As Microsoft.Office.Interop.Excel.Workbook = CType(excelApplication.Workbooks.Add(System.Reflection.Missing.Value), Microsoft.Office.Interop.Excel.Workbook)
        'Dim excelSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(excelWorkbook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        ''add the columns
        'For i As Integer = 0 To DS.Tables(TableName).Columns.Count - 1
        '    CType(excelSheet.Cells(1, i + 1), Microsoft.Office.Interop.Excel.Range).Value2 = DS.Tables(TableName).Columns(i).ColumnName
        'Next i

        ''set the column styles
        'excelSheet.Range(excelSheet.Cells(1, 1), excelSheet.Cells(1, DS.Tables(TableName).Columns.Count)).Font.Bold = True
        'excelSheet.Range(excelSheet.Cells(1, 1), excelSheet.Cells(1, DS.Tables(TableName).Columns.Count)).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        'For i As Integer = 0 To DS.Tables(TableName).Rows.Count - 1
        '    excelSheet.Cells.Range(excelSheet.Cells(i + 2, 1), excelSheet.Cells(i + 2, DS.Tables(TableName).Columns.Count)).NumberFormat = "@"
        '    excelSheet.Cells.Range(excelSheet.Cells(i + 2, 1), excelSheet.Cells(i + 2, DS.Tables(TableName).Columns.Count)).Value2 = DS.Tables(TableName).Rows(i).ItemArray
        'Next i
        'excelApplication.Visible = False
        'excelApplication.DisplayAlerts = False
        'excelWorkbook.Activate()

        'Dim FilePath As String = aspxPage.Server.MapPath("~/DocCRM") & "/DS" & Format(Now, "yyyyMMddHHmmss") & ".xls"

        'If Not System.IO.Directory.Exists(aspxPage.Server.MapPath("~/DocCRM")) Then
        '    System.IO.Directory.CreateDirectory(aspxPage.Server.MapPath("~/DocCRM"))
        'End If
        'excelWorkbook.SaveAs(FilePath)
        'excelWorkbook.Close()
        'excelWorkbook = Nothing
        'excelApplication.Quit()
        'excelApplication = Nothing
        'Dim proc As System.Diagnostics.Process
        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next

        'Dim strHdr As String = ""
        'strHdr = "attachment;filename=ExcelFile.xls"

        'With aspxPage.Response
        '    .Clear()
        '    .ContentType = "application/vnd.ms-excel"
        '    .ContentEncoding = System.Text.Encoding.Default
        '    .AppendHeader("Content-Disposition", strHdr)
        '    .WriteFile(FilePath)
        '    .Flush()
        '    .Clear()
        '    .Close()
        'End With
        'If System.IO.File.Exists(FilePath) Then
        '    System.IO.File.Delete(FilePath)
        'End If

    End Sub

    '------------------------------------------------------------------------------
    'Function to populate combo box. Added By Chan Nyein Zaw(Chan)
    '------------------------------------------------------------------------------
    'Public Shared Sub PopulateCombo(ByRef cmbDD As DropDownList, ByVal sSelectStatement As String, ByVal sTextField As String, ByVal sKeyField As String, ByVal sWhereClause As String, ByVal sBlankLineText As String)
    '    Dim dsTable As DataTable = New DataTable
    '    Dim sBlankLineValue As String = sBlankLineText
    '    Try
    '        With New clsDataAccessBase
    '            Dim sSQL As String = sSelectStatement
    '            If sWhereClause <> "" Then sSQL += " WHERE " + sWhereClause
    '            If .LoadDataBySQL(dsTable, sSQL) Then
    '                If sBlankLineText <> "" Then
    '                    cmbDD.Items.Insert(0, (New ListItem(sBlankLineText, sBlankLineValue)))
    '                    cmbDD.DataSource = dsTable.DataSet
    '                    cmbDD.DataValueField = sKeyField
    '                    cmbDD.DataBind()
    '                End If
    '            End If
    '        End With
    '    Catch ex As Exception

    '    Finally
    '        dsTable.Dispose()
    '    End Try
    'End Sub


    '-----------------------------------------------------------------------------------------
    'This function uses for set focus.
    'Input  : sControlName-> ID Name of Control.
    '       : oLiteral-> Literal object on page that you want to alert.
    'Output :
    '-----------------------------------------------------------------------------------------
    'Public Shared Sub SetFocus(ByVal sControlName As String, ByRef oLiteral As WebControls.Literal)
    '    Dim strjscript As String

    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "var t = document.getElementById('" & sControlName & "');"
    '    strjscript = strjscript & "t.focus();"
    '    strjscript = strjscript & "</script>"
    '    oLiteral.Text = strjscript  'Set the literal control's text to the JScript code
    'End Sub

    '-----------------------------------------------------------------------------------------
    'This function uses for alert message.
    'Input  : sMessage-> Input string such as Error message.
    '       : oLiteral-> Literal object on page that you want to alert.
    'Output :
    '-----------------------------------------------------------------------------------------
    'Public Shared Sub AlertMsgBox(ByVal sMessage As String, ByRef oLiteral As WebControls.Literal)
    '    Dim strjscript As String
    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "window.alert(""" & Replace(Replace(sMessage, """", "'"), vbCrLf, "\n") & """);"
    '    strjscript = strjscript & "</script>"

    '    oLiteral.Text = strjscript  'Set the literal control's text to the JScript code

    'End Sub


    '-----------------------------------------------------------------------------------------
    'This function uses for alert message.
    'Input  : sMessage-> Input string such as Error message.
    '       : oLiteral-> Literal object on page that you want to alert.
    'Output :
    '-----------------------------------------------------------------------------------------

    '  need to rewrite !!

    'Public Shared Sub AlertMsgBox_Return(ByVal sMessage As String, ByRef oLiteral As WebControls.Literal)
    '    Dim strjscript As String
    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "var theFrom = window.opener;"
    '    strjscript = strjscript & "var thisFrom = this;"
    '    strjscript = strjscript & "window.alert(""" & Replace(Replace(sMessage, """", "'"), vbCrLf, "\n") & """);"
    '    strjscript = strjscript & "theFrom.document.Form1.submit();"
    '    strjscript = strjscript & "thisFrom.close();"
    '    strjscript = strjscript & "</script>"

    '    oLiteral.Text = strjscript  'Set the literal control's text to the JScript code

    'End Sub


    'Public Shared Sub Return_Caller(ByRef oLiteral As WebControls.Literal)
    '    Dim strjscript As String
    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "var theFrom = window.opener;"
    '    strjscript = strjscript & "theFrom.document.Form1.submit();"
    '    strjscript = strjscript & "this.close();"
    '    strjscript = strjscript & "</script>"

    '    oLiteral.Text = strjscript  'Set the literal control's text to the JScript code

    'End Sub

    '-----------------------------------------------------------------------------------------
    'This function uses for close window.
    'Input  : sControlName-> ID Name of Control.
    '       : oLiteral-> Literal object on page that you want to alert.
    'Output :
    '-----------------------------------------------------------------------------------------
    'Public Shared Sub CloseWindow(ByRef oLiteral As WebControls.Literal)
    '    Dim strjscript As String

    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "window.close();"
    '    strjscript = strjscript & "</script>"
    '    oLiteral.Text = strjscript  'Set the literal control's text to the JScript code
    'End Sub
    '---------------------------------------------------------------------------
    'Added By Chan Nyein Zaw
    '---------------------------------------------------------------------------
    'Public Shared Function ApplicationPath(ByVal RqAppPath As String) As String
    '    If RqAppPath <> "/" Then
    '        ApplicationPath = RqAppPath
    '    Else
    '        ApplicationPath = ""
    '    End If
    'End Function
    'Public Shared Function AlertMsgYesNoBox(ByVal sMessage As String, ByRef oLiteral As WebControls.Literal) As Boolean
    '    Dim strjscript As String
    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "var returnValue;"
    '    strjscript = strjscript & "returnValue = window.confirm(""" & Replace(Replace(sMessage, """", "'"), vbCrLf, "\n") & """);"
    '    strjscript = strjscript & "this.ReturnHid.value = returnValue"
    '    strjscript = strjscript & "</script>"
    '    oLiteral.Text = strjscript  'Set the literal control's text to the JScript code
    'End Function

    'Public Shared Function PrintMe(ByRef oLiteral As WebControls.Literal) As Boolean
    '    Dim strjscript As String
    '    oLiteral.Text = ""
    '    strjscript = "<script language=""javascript"">"
    '    strjscript = strjscript & "window.print()"
    '    strjscript = strjscript & "</script>"
    '    oLiteral.Text = strjscript
    'End Function


    'Get Schema of table function returning Datatable



    'Public Shared Function GetSchemaofTable(ByVal TableName As String) As DataTable
    '    Dim SchemaTable As DataTable
    '    Dim oConnection As SqlClient.SqlConnection
    '    Dim oCommand As SqlClient.SqlCommand
    '    Dim oReader As SqlClient.SqlDataReader
    '    Try
    '        oConnection = New SqlClient.SqlConnection(clsApplicationConfiguration.ConRSA_Motor)
    '        oCommand = New SqlClient.SqlCommand("SELECT * FROM " & TableName, oConnection)
    '        oConnection.Open()
    '        oReader = oCommand.ExecuteReader(CommandBehavior.SchemaOnly)
    '        SchemaTable = oReader.GetSchemaTable
    '    Catch ex As Exception
    '        Return Nothing
    '    Finally
    '        If Not oReader Is Nothing Then
    '            oReader.Close()
    '        End If
    '        If Not oCommand Is Nothing Then
    '            oCommand.Dispose()
    '        End If
    '        If Not oConnection Is Nothing Then
    '            oConnection.Close()
    '            oConnection.Dispose()
    '        End If
    '    End Try
    '    Return SchemaTable
    'End Function



    Public Shared Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function

    Public Shared Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function



    Public Shared Function AuthenticateUser(domain As String, username As String, password As String, LdapPath As String, ByRef Errmsg As String) As Boolean
        Errmsg = ""
        Dim domainAndUsername As String = domain & "\" & username
        Dim entry As New DirectoryEntry(LdapPath, domainAndUsername, password)
        Try
            ' Bind to the native AdsObject to force authentication.
            Dim obj As [Object] = entry.NativeObject
            Dim search As New DirectorySearcher(entry)
            search.Filter = "(SAMAccountName=" & username & ")"
            search.PropertiesToLoad.Add("cn")
            Dim result As SearchResult = search.FindOne()
            If result Is Nothing Then
                Return False
            End If
            ' Update the new path to the user in the directory
            LdapPath = result.Path
            Dim _filterAttribute As String = DirectCast(result.Properties("cn")(0), [String])
        Catch ex As Exception
            Errmsg = ex.Message
            Return False
            Throw New Exception("Error authenticating user." + ex.Message)
        End Try
        Return True
    End Function
End Class

