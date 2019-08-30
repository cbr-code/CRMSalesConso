
Imports System.Data.SqlClient


Public Class SQLControl
    'Dim dbname As String = "cindysba_sales"
    'Dim dbhost As String = "cindysbakery.com"
    'Dim user As String = "cindysba_root"
    'Dim pass As String = "csd@cindys"

    'Private DBCon As New SqlConnection("SERVER=csdsvr;DATABASE=be;Integrated Security=SSPI;")
    'Private DBCon2 As New SqlConnection("SERVER=csdsvr;DATABASE=cindys;Integrated Security=SSPI;")

    'MySQL on Live // It works.. God only knows why...
    'Private DBCon1 As New MySqlConnection(String.Format("server={0}; user id={1}; password={2}; database={3}; pooling=false", dbhost, user, pass, dbname))

    'MS SQL HRSYS
    'Private DBCon2 As New SqlConnection("Server=intranet;Database=cbakerpaydb;User=csd;Pwd=P@$$w0rd;")

    'Main BE LIve
    Public DBCon3 As New SqlConnection("Server=172.16.0.12;Database=SALESINV;User=sa;Pwd=Passw0rd;")
    'Public DBCon3 As New SqlConnection("SERVER=(local);DATABASE=mainbe;Integrated Security=SSPI;")

    'Main BE Staging
    'Private DBCon3 As New SqlConnection("Server=csdsvr;Database=mainbe_staging;User=sa;Pwd=Passw0rd;")

    'OSMS
    'Public DBCon4 As New SqlConnection("Server=solomon_cbc;Database=osms;User=sa;Pwd=Passw0rd;")

    'Local/Cindys
    'Private DBCon5 As New SqlConnection("SERVER=(local);DATABASE=cindy;Integrated Security=SSPI;")


    'MS SQL
    Private DBCmd As SqlCommand

    'Mysql
    'Private MyDBCmd As MySqlCommand

    ' DB DATA For MS SQL
    Public DBDA As SqlDataAdapter
    Public DBDT As DataTable

    ' DB DATA For MY SQL
    'Public MyDBDA As MySqlDataAdapter
    Public MyDBDT As DataTable

    ' DB Reader MY SQL
    'Public MyDBDR As MySqlDataReader
    Public MyDBS As DataSet


    ' DB Reader MS SQL
    Public DBDR As SqlDataReader
    Public DBS As DataSet


    ' QUERY PARAMETERS MS SQL
    Public Params As New List(Of SqlParameter)

    ' QUERY PARAMETERS MY SQL
    'Public MyParams As New List(Of MySqlParameter)

    ' QUERY STATISTICS 
    Public RecordCount As Integer
    Public Exception As String
    Public LastID As Integer

    Public Sub New()
    End Sub

    ' ALLOW CONNECTION STRING OVERRIDE
    Public Sub New(ConnectionString As String)

        'DBCon1 = New MySqlConnection(ConnectionString)
        'DBCon2 = New SqlConnection(ConnectionString)
        DBCon3 = New SqlConnection(ConnectionString)

    End Sub

    ' EXECUTE QUERY SUB For BE DB
    'Public Sub ExecQuery(Query As String)
    '    ' RESET QUERY STATS
    '    RecordCount = 0
    '    Exception = ""

    '    Try

    '        DBCon2.Open()

    '        ' CREATE DB COMMAND
    '        DBCmd = New SqlCommand(Query, DBCon2)

    '        ' LOAD PARAMS INTO DB COMMAND
    '        Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

    '        ' CLEAR PARAM LIST
    '        Params.Clear()

    '        ' EXECUTE COMMAND AND FILL DATASET
    '        DBDT = New DataTable

    '        DBDA = New SqlDataAdapter(Query, DBCon2)
    '        DBS = New DataSet
    '        DBDA.Fill(DBS)

    '        RecordCount = DBDA.Fill(DBDT)

    '        'RecordCount = DBDT.Rows.Count

    '    Catch ex As Exception
    '        ' CAPTURE ERRORS
    '        Exception = "ExecQuery Error: " & vbNewLine & ex.Message
    '    End Try

    '    ' CLOSE CONNECTION

    '    If DBCon2.State = ConnectionState.Open Then DBCon2.Close()


    'End Sub

    'Public Sub LiveExecQuery(Query As String)
    '    ' RESET QUERY STATS
    '    RecordCount = 0
    '    Exception = ""

    '    Try

    '        DBCon1.Open()

    '        ' CREATE DB COMMAND
    '        MyDBCmd = New MySqlCommand(Query, DBCon1)

    '        ' LOAD PARAMS INTO DB COMMAND
    '        MyParams.ForEach(Sub(p) MyDBCmd.Parameters.Add(p))

    '        ' CLEAR PARAM LIST
    '        MyParams.Clear()

    '        ' EXECUTE COMMAND AND FILL DATASET
    '        MyDBDT = New DataTable

    '        MyDBDA = New MySqlDataAdapter(Query, DBCon1)
    '        MyDBS = New DataSet
    '        ' RecordCount = MyDBDA.Fill(MyDBS)

    '        RecordCount = MyDBDA.Fill(MyDBDT)



    '    Catch ex As Exception
    '        ' CAPTURE ERRORS
    '        Exception = "MyExecQuery Error: " & vbNewLine & ex.Message
    '    End Try

    '    ' CLOSE CONNECTION

    '    If DBCon1.State = ConnectionState.Open Then DBCon1.Close()


    'End Sub
    'Public Sub ExecProc(SPQuery As String)
    '    ' RESET QUERY STATS
    '    RecordCount = 0
    '    Exception = ""

    '    Try
    '        DBCon2.Open()
    '        ' CREATE DB COMMAND
    '        DBCmd = New SqlCommand(SPQuery, DBCon2)

    '        DBCmd.CommandType = CommandType.StoredProcedure

    '        ' LOAD PARAMS INTO DB COMMAND
    '        Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

    '        ' CLEAR PARAM LIST
    '        Params.Clear()

    '        ' EXECUTE COMMAND AND FILL DATASET
    '        DBDT = New DataTable
    '        DBDA = New SqlDataAdapter(DBCmd)
    '        RecordCount = DBDA.Fill(DBDT)
    '        'RecordCount = DBDT.Rows.Count

    '    Catch ex As Exception
    '        ' CAPTURE ERRORS
    '        Exception = "ExecParameter Error: " & vbNewLine & ex.Message
    '    End Try

    '    ' CLOSE CONNECTION
    '    If DBCon2.State = ConnectionState.Open Then DBCon2.Close()
    'End Sub

    ' EXECUTE QUERY SUB For Main DB
    Public Sub MDBExecQuery(Query As String)
        ' RESET QUERY STATS

        RecordCount = 0
        Exception = ""

        Try
            DBCon3.Open()

            ' CREATE DB COMMAND
            DBCmd = New SqlCommand(Query, DBCon3)

            ' LOAD PARAMS INTO DB COMMAND
            Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

            ' CLEAR PARAM LIST
            Params.Clear()

            ' EXECUTE COMMAND AND FILL DATASET
            DBDT = New DataTable

            DBDA = New SqlDataAdapter(Query, DBCon3)
            DBS = New DataSet

            ' DBDA.Fill(DBS)

            DBDA.Fill(DBDT)

            RecordCount = DBDT.Rows.Count

        Catch ex As Exception
            ' CAPTURE ERRORS
            Exception = "MyExecQuery Error: " & vbNewLine & ex.Message
        End Try

        ' CLOSE CONNECTION
        If DBCon3.State = ConnectionState.Open Then DBCon3.Close()

    End Sub

    Public Sub MDBExecProc(SPQuery As String)
        ' RESET QUERY STATS
        RecordCount = 0
        Exception = ""

        Try
            DBCon3.Open()
            ' CREATE DB COMMAND
            DBCmd = New SqlCommand(SPQuery, DBCon3)

            DBCmd.CommandType = CommandType.StoredProcedure

            ' LOAD PARAMS INTO DB COMMAND
            Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

            ' CLEAR PARAM LIST
            Params.Clear()

            ' EXECUTE COMMAND AND FILL DATASET
            DBDT = New DataTable
            DBDA = New SqlDataAdapter(DBCmd)

            RecordCount = DBDA.Fill(DBDT)
            'RecordCount = DBDT.Rows.Count

        Catch ex As Exception
            ' CAPTURE ERRORS
            Exception = "MProcExecParameter Error: " & vbNewLine & ex.Message
        End Try

        ' CLOSE CONNECTION
        If DBCon3.State = ConnectionState.Open Then DBCon3.Close()
    End Sub


    Public Sub MDBExecProcGetId(SPQuery As String)
        ' RESET QUERY STATS
        RecordCount = 0
        Exception = ""
        LastID = 0

        Try
            DBCon3.Open()
            ' CREATE DB COMMAND
            DBCmd = New SqlCommand(SPQuery, DBCon3)

            DBCmd.CommandType = CommandType.StoredProcedure

            ' LOAD PARAMS INTO DB COMMAND
            Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

            ' CLEAR PARAM LIST
            Params.Clear()

            ' EXECUTE COMMAND AND FILL DATASET
            DBDT = New DataTable
            DBDA = New SqlDataAdapter(DBCmd)

            LastID = Convert.ToInt32(DBCmd.ExecuteScalar())

            'RecordCount = DBDA.Fill(DBDT)
            'RecordCount = DBDT.Rows.Count

        Catch ex As Exception
            ' CAPTURE ERRORS
            Exception = "MProcExecParameter Error: " & vbNewLine & ex.Message
        End Try

        ' CLOSE CONNECTION
        If DBCon3.State = ConnectionState.Open Then DBCon3.Close()
    End Sub

    ' ADD PARAM SUB
    Public Sub AddParam(Name As String, Value As Object)
        Dim NewParam As New SqlParameter(Name, Value)
        Params.Add(NewParam)
    End Sub


    'Public Sub AddMyParam(Name As String, Value As Object)
    '    Dim NewMyParam As New MySqlParameter(Name, Value)
    '    MyParams.Add(NewMyParam)
    'End Sub


    ' EXECUTE QUERY SUB For OSMS DB
    'Public Sub ExecQryOSMS(Query As String)
    '    ' RESET QUERY STATS
    '    RecordCount = 0
    '    Exception = ""

    '    Try

    '        DBCon4.Open()

    '        ' CREATE DB COMMAND
    '        DBCmd = New SqlCommand(Query, DBCon4)

    '        ' LOAD PARAMS INTO DB COMMAND
    '        Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

    '        ' CLEAR PARAM LIST
    '        Params.Clear()

    '        ' EXECUTE COMMAND AND FILL DATASET
    '        DBDT = New DataTable

    '        DBDA = New SqlDataAdapter(Query, DBCon4)
    '        DBS = New DataSet
    '        DBDA.Fill(DBS)

    '        RecordCount = DBDA.Fill(DBDT)

    '        'RecordCount = DBDT.Rows.Count

    '    Catch ex As Exception
    '        ' CAPTURE ERRORS
    '        Exception = "ExecQuery Error: " & vbNewLine & ex.Message
    '    End Try

    '    ' CLOSE CONNECTION

    '    If DBCon4.State = ConnectionState.Open Then DBCon4.Close()


    'End Sub

    '' EXECUTE QUERY SUB For OSMS DB
    'Public Sub ExecQryLocal(Query As String)
    '    ' RESET QUERY STATS
    '    RecordCount = 0
    '    Exception = ""

    '    Try

    '        DBCon5.Open()

    '        ' CREATE DB COMMAND
    '        DBCmd = New SqlCommand(Query, DBCon5)

    '        ' LOAD PARAMS INTO DB COMMAND
    '        Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

    '        ' CLEAR PARAM LIST
    '        Params.Clear()

    '        ' EXECUTE COMMAND AND FILL DATASET
    '        DBDT = New DataTable

    '        DBDA = New SqlDataAdapter(Query, DBCon5)
    '        DBS = New DataSet
    '        DBDA.Fill(DBS)

    '        RecordCount = DBDA.Fill(DBDT)

    '        'RecordCount = DBDT.Rows.Count

    '    Catch ex As Exception
    '        ' CAPTURE ERRORS
    '        Exception = "ExecQuery Error: " & vbNewLine & ex.Message
    '    End Try

    '    ' CLOSE CONNECTION

    '    If DBCon5.State = ConnectionState.Open Then DBCon5.Close()


    'End Sub

End Class
