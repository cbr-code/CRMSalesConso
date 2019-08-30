Imports System.Xml
Imports System.Text.RegularExpressions
Module AppModule

    Public now As DateTime = DateTime.Now
    Public uid As Integer = 0
    Public uname As String = ""
    Public ufname As String = ""
    Public usertype As Integer = 0
    Public uemail As String = ""
    Dim SQL As New SQLControl
    Public svrtime As String = ""
    Public Function NotEmpty(val As String) As Boolean
        If Not String.IsNullOrEmpty(val) Then Return True Else Return False
    End Function

    Public Sub createUploadSettings(ByVal StoreCode As String, ByVal StoreName As String, ByVal User As String, ByVal DateTran As String, ByVal FileName As String, ByVal DateTime As String, ByVal Manual As String, ByVal Email As String, ByVal writer As XmlTextWriter)

        writer.WriteStartElement("UploadSettings")
        writer.WriteStartElement("StoreCode")
        writer.WriteString(StoreCode)
        writer.WriteEndElement()

        writer.WriteStartElement("StoreName")
        writer.WriteString(StoreName)
        writer.WriteEndElement()

        writer.WriteStartElement("User")
        writer.WriteString(User)
        writer.WriteEndElement()

        writer.WriteStartElement("DateTran")
        writer.WriteString(DateTran)
        writer.WriteEndElement()

        writer.WriteStartElement("FileName")
        writer.WriteString(FileName)
        writer.WriteEndElement()

        writer.WriteStartElement("DateTime")
        writer.WriteString(DateTime)
        writer.WriteEndElement()

        writer.WriteStartElement("UploadType")
        writer.WriteString(Manual)
        writer.WriteEndElement()

        writer.WriteStartElement("Email")
        writer.WriteString(Email)
        writer.WriteEndElement()

        writer.WriteEndElement()
    End Sub


    Public Function validateEmail(emailAddress) As Boolean

        ' Dim email As New Regex("^(?<user>[^@]+)@(?<host>.+)$")
        Dim email As New Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")
        If email.IsMatch(emailAddress) Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function CheckFileExists(ByVal Files As IEnumerable(Of String)) As IEnumerable(Of String)

        Dim sb As New System.Text.StringBuilder

        For Each File As String In Files

            If Not IO.File.Exists(File) Then
                sb.AppendLine(File)
            End If

        Next File

        Return sb.ToString.Split({Environment.NewLine},
                                 StringSplitOptions.RemoveEmptyEntries)

    End Function

    Public Function IsConnectionAvailable() As Boolean
        Dim objUrl As New System.Uri("http://google.com/")
        Dim objWebReq As System.Net.WebRequest
        objWebReq = System.Net.WebRequest.Create(objUrl)
        Dim objresp As System.Net.WebResponse

        Try
            objresp = objWebReq.GetResponse
            objresp.Close()
            objresp = Nothing
            Return True

        Catch ex As Exception
            objresp = Nothing
            objWebReq = Nothing
            Return False
        End Try
    End Function

    Public Function GetMDB()

        Try
            Sql.DBCon3.Open()
            Sql.DBCon3.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


    Function getServerTime()



        ' EXECUTE QUERY
        SQL.MDBExecQuery("SELECT GETDATE() as gdate ")

        ' REPORT & ABORT ON ERRORS
        'If NotEmpty(SQLMDB.Exception) Then MsgBox(SQL.Exception) : Exit Sub

        ' POPULATE THREAD LIST
        For Each r As DataRow In SQL.DBDT.Rows
            svrtime = r("gdate")
        Next

        Return svrtime


    End Function


    Function GetFilename(Orig_FileName)

        SQL.MDBExecQuery("select file_name from Settings where file_name = '" & Orig_FileName & "'")

        If SQL.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Function GetTmpTbl() As DataTable
        ' Create new DataTable instance.
        Dim table As New DataTable

        table.Columns.Add("TranDate", GetType(DateTime))


        Return table
    End Function

End Module
