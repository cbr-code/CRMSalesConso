Imports System.IO
Imports System.Net
Imports System.Xml
Imports Ionic.Zip
Imports Microsoft.Office

Public Class cscParser
    Dim SQL As New SQLControl
    Dim SQLMDB As New SQLControl
    Dim xmlTranID As String, StoreCode As String = "", StoreName As String = "", User As String = "", DateTran As String = "", FileName As String = "", DateTime As String = "", UploadType As String, Email As String, strplucode As String
    Dim GetSettingsID, GetSettingsInvID, gTranID, second As Integer
    Dim getFName, getOName, getFId, Orig_FileName, getUploadStamp, ExtFileName, msgstatus, getEmail, getEmpname, DTranDay, getCSVFininsedSales, getCSVTrans, getCSVPayments As String
    Dim getFolderXml As String = "C:\\GetExportedLocal\"
    Dim getFolderZip As String = "C:\inetpub\CBRServices\UploadedFiles\"
    Dim PubPercent As Decimal
    Dim getCRMFolderXml As String = "C:\\GetCRMExported\"
    Dim TranDateSales As DateTime
    Dim boolDc, bchk As Boolean

    Private Sub cscParser_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(8000, "CRM Sales Conso Trans", "Working on background", ToolTipIcon.Info)
        End If
    End Sub



    Private Sub cscParser_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call Load_Csv()

        Application.Exit()

    End Sub



    Public Sub Load_Csv()


        Me.WindowState = FormWindowState.Minimized
        'Check server conection
        If Not GetMDB() Then
            'rst5mins()
            Application.Exit()
            Exit Sub
        End If

        'Check connections 
        If Not IsConnectionAvailable() Then
            'rst5mins()
            'Application.Exit()
            Exit Sub
        End If

        'Check Directory
        If (Not System.IO.Directory.Exists(getCRMFolderXml)) Then
            System.IO.Directory.CreateDirectory(getCRMFolderXml)
        End If

        Call deleteExFilesX(getCRMFolderXml)

        'Get Extracted files
        boolDc = True

        'Dim ftpRequest As FtpWebRequest = DirectCast(WebRequest.Create("ftp://cindysbakery.com/"), FtpWebRequest)
        'ftpRequest.Credentials = New NetworkCredential("crm@cindysbakery.com", "csd@cindys")
        'ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory
        'Dim response As FtpWebResponse = DirectCast(ftpRequest.GetResponse(), FtpWebResponse)
        'Dim streamReader As New StreamReader(response.GetResponseStream())
        'Dim directories As New List(Of String)()

        'Dim line As String = streamReader.ReadLine()
        'While Not String.IsNullOrEmpty(line)
        '    directories.Add(line)
        '    line = streamReader.ReadLine()
        'End While
        'streamReader.Close()

        'Using ftpClient As New WebClient()
        '    ftpClient.Credentials = New System.Net.NetworkCredential("crm@cindysbakery.com", "csd@cindys")

        '    For i As Integer = 0 To directories.Count - 1
        '        If directories(i).Contains("CRM_Sales.zip") Then

        '            Dim path As String = "ftp://cindysbakery.com/" + directories(i).ToString()
        '            Dim trnsfrpth As String = "C:\\GetCRMExported\" + directories(i).ToString()
        '            Dim trnsfrpth2 As String = "C:\\BackupZipFiles\" + directories(i).ToString()

        '            Try
        '                ftpClient.DownloadFile(path, trnsfrpth)
        '                ftpClient.DownloadFile(path, trnsfrpth2)
        '                getOName = directories(i).ToString()
        '                getFName = directories(i).ToString()
        '                Exit For
        '            Catch
        '                Exit Sub
        '            End Try

        '        End If
        '    Next
        'End Using


        Dim Folder As New IO.DirectoryInfo(getFolderZip)
        For Each File As IO.FileInfo In Folder.GetFiles("*SalesCRM.zip", IO.SearchOption.TopDirectoryOnly)
            getOName = File.Name
            Exit For
        Next


        'Check file
        If IsNothing(getOName) Then
            Exit Sub
        End If


        ' get Zip Backup
        Dim zipSource As String = Path.Combine(getFolderZip, getOName)
        Dim zipBackCopy As String = Path.Combine("C:\\BackupZipFiles\", getOName)
        File.Copy(zipSource, zipBackCopy, True)


        'Extract File
        Call ExFile()

        'Validate XML Settings Upload 
        If GetSettingsCRM() Then
            Exit Sub
        End If

        Call SetCSVData()
        Call GetFinishedSales()
        Call GetTransactions()
        Call GetPayments()

        'msgstatus = "Successfully Transmitted Sales Data to the Main Server at " & getServerTime()

        Call setEmailSend(getOName, "", getEmail, "", "", "CRM Consolidator", "Succesfully Transmitted", 5)

        Call deleteExFilesX(getCRMFolderXml)

        Call deleteExFiles(getFolderZip)



        'Call deleteFtpZipFile(getOName)

        frmMain.NotifyIcon1.Visible = True

        frmMain.NotifyIcon1.ShowBalloonTip(8000, "CRM Consolidator", "Successfully Transmitted Sales Data: " & StoreName & " - " & getOName, ToolTipIcon.Info)

        'getOName = Nothing

        'rst10sec()

        'Application.Exit()

    End Sub



    ' Dim TranDateSales As DateTime = Convert.ToDateTime(vDate2)


    Public Sub deleteFtpZipFile(Orig_FileName As String)
        Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://ftp.cindysbakery.com/" + Orig_FileName), System.Net.FtpWebRequest)
        request.Credentials = New System.Net.NetworkCredential("crm@cindysbakery.com", "csd@cindys")
        request.Method = System.Net.WebRequestMethods.Ftp.DeleteFile
        request.GetResponse()
    End Sub

    Public Sub SetCSVData()

        Dim fSettingsXml As String() = Directory.GetFiles(getCRMFolderXml, "a.xml")
        For Each FileName In fSettingsXml
            If (System.IO.File.Exists(FileName)) Then
                'Read Settings 
                If FileName = getCRMFolderXml & "" & "a.xml" Then

                    'Insert Settings 
                    ReadXmlCRMSettings(FileName)

                End If
            End If
        Next


        'Finishsales
        Dim fFinishedSales As String() = Directory.GetFiles(getCRMFolderXml, "Z004_" & DTranDay & ".CSV")
        For Each FileName In fFinishedSales
            If (System.IO.File.Exists(FileName)) Then
                'Read Settings 
                If FileName = getCSVFininsedSales Then

                    getFinishedSales(FileName)

                End If
            End If
        Next

        'Transaction 
        Dim fTransaction As String() = Directory.GetFiles(getCRMFolderXml, "Z002_" & DTranDay & ".CSV")
        For Each FileName In fTransaction
            If (System.IO.File.Exists(FileName)) Then
                'Read Settings 
                If FileName = getCSVTrans Then

                    getTransaction(FileName)

                End If
            End If
        Next


        'Payments 
        Dim fPayments As String() = Directory.GetFiles(getCRMFolderXml, "Z001_" & DTranDay & ".CSV")
        For Each FileName In fPayments
            If (System.IO.File.Exists(FileName)) Then
                'Read Settings 
                If FileName = getCSVPayments Then

                    getPayments(FileName)

                End If
            End If
        Next



    End Sub


    Public Sub ReadXmlCRMSettings(xmlFileName As String)

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlFileName)
        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/Settings")
        For Each node As XmlNode In nodes

            'StoreCode = node.SelectSingleNode("StoreCode").InnerText

            SQLMDB.AddParam("@tranid", 0)
            SQLMDB.AddParam("@st_code", node.SelectSingleNode("StoreCode").InnerText)
            SQLMDB.AddParam("@st_name", node.SelectSingleNode("StoreName").InnerText)
            SQLMDB.AddParam("@user_file", node.SelectSingleNode("User").InnerText)
            SQLMDB.AddParam("@date_tran", Convert.ToDateTime(node.SelectSingleNode("DateTran").InnerText))
            SQLMDB.AddParam("@file_name", node.SelectSingleNode("FileName").InnerText)
            SQLMDB.AddParam("@ext_date", Convert.ToDateTime(node.SelectSingleNode("DateTime").InnerText))
            SQLMDB.AddParam("@date_wupload", Convert.ToDateTime(node.SelectSingleNode("DateTime").InnerText))
            SQLMDB.AddParam("@date_mupload", now)
            SQLMDB.AddParam("@upload_type", node.SelectSingleNode("UploadType").InnerText)
            SQLMDB.AddParam("@email", node.SelectSingleNode("Email").InnerText)
            SQLMDB.AddParam("@Mode", "ADD")
            SQLMDB.MDBExecProcGetId("spm_settings")


            GetSettingsID = SQLMDB.LastID

            If NotEmpty(SQLMDB.Exception) Then MsgBox(SQLMDB.Exception) : Exit Sub


        Next

    End Sub


    Public Sub GetFinishedSales()
        Dim keypos As Integer
        Dim vDate As String
        Dim vDate2 As String
        Dim vTime As String
        'Dim TranDate As Date
        keypos = 0
        For Each row As DataGridViewRow In dgvSales.Rows
            If Not (row.Cells(0).Value = Nothing) Then

                'Get DateTime
                If keypos = 6 Then
                    vDate = row.Cells(1).Value
                End If

                If keypos = 7 Then
                    vTime = row.Cells(1).Value
                    vDate2 = vDate + " " + vTime
                    TranDateSales = Convert.ToDateTime(vDate2)
                End If

                If keypos > 8 Then

                    If (row.Cells(1).Value <> "") Then
                        If (row.Cells(2).Value <> "0") Then
                            'Insert

                            'SQL.MDBExecQuery("SELECT * FROM VCRM where plucode =  '" & row.Cells(0).Value & "' ")

                            strplucode = row.Cells(0).Value.TrimStart("0"c)

                            SQL.MDBExecQuery("SELECT * FROM vwPosProduct where FieldStyleCode =  '" & strplucode & "' and scode = '" & StoreCode & "'")

                            'If SQL.RecordCount = 0 Then

                            'MsgBox("ok")

                            'End If

                            For Each r As DataRow In SQL.DBDT.Rows
                                'Insert FinishedSales
                                SQLMDB.AddParam("@SID", GetSettingsID)
                                SQLMDB.AddParam("@LineID", 0)
                                SQLMDB.AddParam("@TransactionNo", 0)
                                SQLMDB.AddParam("@ProductID", r("ProductID"))

                                Try
                                    SQLMDB.AddParam("@ProductCode", r("ProductCode"))
                                Catch
                                    SQLMDB.AddParam("@ProductCode", vbNullString)
                                End Try

                                SQLMDB.AddParam("@Barcode", r("ProductCode"))


                                If r("ProductCode") = "CKS100" Then
                                    SQLMDB.AddParam("@Description", "CAKELESS")
                                ElseIf r("ProductCode") = "CKS099" Then
                                    SQLMDB.AddParam("@Description", "CAKE DEPOSIT")
                                Else
                                    SQLMDB.AddParam("@Description", r("Description"))
                                End If

                                Try
                                    SQLMDB.AddParam("@UOM", r("uom"))
                                Catch
                                    SQLMDB.AddParam("@UOM", vbNullString)
                                End Try

                                SQLMDB.AddParam("@Qty", row.Cells(2).Value)
                                SQLMDB.AddParam("@Packing", 1)
                                SQLMDB.AddParam("@TotalQty", row.Cells(2).Value)

                                SQLMDB.AddParam("@AverageUnitCost", 0)

                                SQLMDB.AddParam("@Price", r("srp"))

                                SQLMDB.AddParam("@Discount", 0)


                                SQLMDB.AddParam("@Allowance", 0)
                                SQLMDB.AddParam("@AmountDiscounted", 0)
                                SQLMDB.AddParam("@ChargeDiscount", 0)
                                SQLMDB.AddParam("@ChargeAllowance", 0)
                                SQLMDB.AddParam("@ChargeAmountDiscounted", 0)
                                SQLMDB.AddParam("@DiscountedPrice", 0)
                                SQLMDB.AddParam("@DiscountDescription", vbNullString)


                                SQLMDB.AddParam("@Extended", Convert.ToDouble(row.Cells(3).Value))
                                SQLMDB.AddParam("@ExtendedDescription", row.Cells(3).Value)

                                SQLMDB.AddParam("@Multiplier", 1)
                                SQLMDB.AddParam("@PriceModeCode", "R")


                                'SQLMDB.AddParam("@xReturn", 0)

                                SQLMDB.AddParam("@ReturnDescription", vbNullString)

                                SQLMDB.AddParam("@ReturnRemarks", vbNullString)

                                SQLMDB.AddParam("@OldTransactionNo", 0)

                                SQLMDB.AddParam("@OldTransactionDate", DateTran)

                                SQLMDB.AddParam("@OldTransactionDiscount", 0)

                                SQLMDB.AddParam("@OldTerminalNo", vbNullString)


                                'SQLMDB.AddParam("@Shift", 0)
                                'SQLMDB.AddParam("@UserID", 0)
                                'SQLMDB.AddParam("@TerminalNo", 0)
                                'SQLMDB.AddParam("@BranchCode", StoreCode)
                                SQLMDB.AddParam("@LogDate", DateTran)
                                SQLMDB.AddParam("@DateTime", DateTran)
                                'SQLMDB.AddParam("@Voided", 0)
                                'SQLMDB.AddParam("@MustReachForFree", 0)
                                'SQLMDB.AddParam("@Points", 0)
                                'SQLMDB.AddParam("@PointsPosted", 0)
                                'SQLMDB.AddParam("@AmountSaved", 0)
                                'SQLMDB.AddParam("@QtyReturned", 0)
                                'SQLMDB.AddParam("@PriceOverride", 0)
                                'SQLMDB.AddParam("@MarkDown", 0)


                                'SQLMDB.AddParam("@SerialNo", vbNullString)


                                'SQLMDB.AddParam("@PROMOPERSONID", vbNullString)

                                'SQLMDB.AddParam("@SONumber", vbNullString)

                                'SQLMDB.AddParam("@TimeScanned", 0)
                                'SQLMDB.AddParam("@Layaway", 0)
                                'SQLMDB.AddParam("@LayawayNumber", 0)
                                'SQLMDB.AddParam("@pVatable", 1)
                                'SQLMDB.AddParam("@pVatPercent", 12)
                                'SQLMDB.AddParam("@ChilledCharge", 0)

                                'SQLMDB.AddParam("@Remarks", vbNullString)

                                'SQLMDB.AddParam("@Senior", vbNullString)
                                'SQLMDB.AddParam("@discounttype1", 0)
                                'SQLMDB.AddParam("@discounttype2", 0)
                                'SQLMDB.AddParam("@discounttype3", 0)
                                'SQLMDB.AddParam("@discounttype4", 0)
                                'SQLMDB.AddParam("@discounttype5", 0)
                                'SQLMDB.AddParam("@discounttype6", 0)
                                'SQLMDB.AddParam("@discounttype7", 0)
                                'SQLMDB.AddParam("@discounttype8", 0)
                                'SQLMDB.AddParam("@discounttype9", 0)
                                'SQLMDB.AddParam("@discounttype10", 0)
                                'SQLMDB.AddParam("@Diplomat", 0)
                                'SQLMDB.AddParam("@BonusPoints", 0)
                                'SQLMDB.AddParam("@Tax", 0)
                                'SQLMDB.AddParam("@TaxID", 0)

                                'SQLMDB.AddParam("@DateTimeStart", vbNullString)

                                'SQLMDB.AddParam("@DateTimeEnd", vbNullString)

                                'SQLMDB.AddParam("@ServiceCharge", 0)

                                'SQLMDB.AddParam("@PriceOverrideReason", vbNullString)

                                'SQLMDB.AddParam("@Discountcode", vbNullString)

                                'SQLMDB.AddParam("@RegisterNo", 0)

                                Try
                                    SQLMDB.AddParam("@FieldACode", r("FieldACode"))
                                Catch
                                    SQLMDB.AddParam("@FieldACode", 0)
                                End Try

                                Try
                                    SQLMDB.AddParam("@FieldBCode", r("FieldBCode"))
                                Catch
                                    SQLMDB.AddParam("@FieldBCode", 0)
                                End Try

                                Try
                                    SQLMDB.AddParam("@FieldCCode", r("FieldCCode"))
                                Catch
                                    SQLMDB.AddParam("@FieldCCode", 0)
                                End Try

                                SQLMDB.AddParam("@UploadStamp", now)
                                SQLMDB.AddParam("@Mode", "ADD")
                                SQLMDB.MDBExecProc("spm_crm_fsales")

                                If NotEmpty(SQLMDB.Exception) Then MsgBox(SQLMDB.Exception) : Exit Sub


                            Next
                        End If
                    End If

                End If


            End If
            keypos = keypos + 1
        Next
    End Sub

    Public Sub GetTransactions()
        Dim keypos As Integer
        Dim vDate As String
        Dim vDate2 As String
        Dim vTime As String

        Dim vMode, vZ, ZCounter As String

        'Dim TranDate As Date
        keypos = 0
        For Each row As DataGridViewRow In dgvTrans.Rows
            If Not (row.Cells(0).Value = Nothing) Then

                'Get DateTime
                If keypos = 4 Then
                    vMode = row.Cells(1).Value
                End If

                If keypos = 5 Then
                    vZ = row.Cells(1).Value
                    ZCounter = vMode + "-" + vZ
                End If


                If keypos = 6 Then
                    vDate = row.Cells(1).Value
                End If

                If keypos = 7 Then
                    vTime = row.Cells(1).Value
                    vDate2 = vDate + " " + vTime
                    TranDateSales = Convert.ToDateTime(vDate2)
                End If

                If keypos > 8 Then

                    If (row.Cells(1).Value <> "") Then
                        'If (row.Cells(2).Value <> "0") Then

                        'Insert Transactions
                        SQLMDB.AddParam("@SID", GetSettingsID)
                            SQLMDB.AddParam("@TransactionNo", vZ)
                            SQLMDB.AddParam("@Mode", ZCounter)
                            SQLMDB.AddParam("@TranDate", TranDateSales)
                            SQLMDB.AddParam("@RecID", row.Cells(0).Value)
                            SQLMDB.AddParam("@TranDesc", row.Cells(1).Value)
                        SQLMDB.AddParam("@Qty", row.Cells(2).Value)

                        Try
                            SQLMDB.AddParam("@Amount", Convert.ToDouble(row.Cells(3).Value))
                        Catch
                            SQLMDB.AddParam("@Amount", 0.00)
                        End Try



                        SQLMDB.AddParam("@UploadStamp", now)
                            SQLMDB.AddParam("@Stype", "Tran")
                            SQLMDB.AddParam("@Action", "ADD")
                            SQLMDB.MDBExecProc("sp_CRMTranPay")

                            If NotEmpty(SQLMDB.Exception) Then MsgBox(SQLMDB.Exception) : Exit Sub

                        'End If
                    End If

                End If


            End If
            keypos = keypos + 1
        Next
    End Sub


    Public Sub GetPayments()
        Dim keypos As Integer
        Dim vDate As String
        Dim vDate2 As String
        Dim vTime As String

        Dim vMode, vZ, ZCounter As String

        'Dim TranDate As Date
        keypos = 0
        For Each row As DataGridViewRow In dgvpayments.Rows
            If Not (row.Cells(0).Value = Nothing) Then

                'Get DateTime
                If keypos = 4 Then
                    vMode = row.Cells(1).Value
                End If

                If keypos = 5 Then
                    vZ = row.Cells(1).Value
                    ZCounter = vMode + "-" + vZ
                End If


                If keypos = 6 Then
                    vDate = row.Cells(1).Value
                End If

                If keypos = 7 Then
                    vTime = row.Cells(1).Value
                    vDate2 = vDate + " " + vTime
                    TranDateSales = Convert.ToDateTime(vDate2)
                End If

                If keypos > 8 Then

                    If (row.Cells(1).Value <> "") Then
                        ' If (row.Cells(2).Value <> "0") Then

                        'Insert Transactions
                        SQLMDB.AddParam("@SID", GetSettingsID)
                            SQLMDB.AddParam("@TransactionNo", vZ)
                            SQLMDB.AddParam("@Mode", ZCounter)
                            SQLMDB.AddParam("@TranDate", TranDateSales)
                            SQLMDB.AddParam("@RecID", row.Cells(0).Value)
                            SQLMDB.AddParam("@TranDesc", row.Cells(1).Value)
                        SQLMDB.AddParam("@Qty", row.Cells(2).Value)

                        Try
                            SQLMDB.AddParam("@Amount", Convert.ToDouble(row.Cells(3).Value))
                        Catch
                            SQLMDB.AddParam("@Amount", 0.00)
                        End Try

                        SQLMDB.AddParam("@UploadStamp", now)
                            SQLMDB.AddParam("@Stype", "Pay")
                            SQLMDB.AddParam("@Action", "ADD")
                            SQLMDB.MDBExecProc("sp_CRMTranPay")

                            If NotEmpty(SQLMDB.Exception) Then MsgBox(SQLMDB.Exception) : Exit Sub

                        ' End If
                    End If

                End If


            End If
            keypos = keypos + 1
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim keypos As Integer
        Dim vDate As String
        Dim vDate2 As String
        Dim vTime As String
        'Dim TranDate As Date
        keypos = 0
        For Each row As DataGridViewRow In dgvSales.Rows
            If Not (row.Cells(0).Value = Nothing) Then


                'Get Date

                If keypos = 6 Then
                    vDate = row.Cells(1).Value
                End If

                If keypos = 7 Then
                    vTime = row.Cells(1).Value
                    vDate2 = vDate + " " + vTime
                    Dim TranDate As DateTime = Convert.ToDateTime(vDate2)
                End If



                If keypos > 9 Then

                    ' MessageBox.Show(row.Cells(0).Value)

                    If (row.Cells(1).Value <> "") Then

                        'MessageBox.Show(row.Cells(1).Value)
                        'insert


                    End If

                End If



                '  If row.Cells(0).Value = "RECORD" Then

                ' MessageBox.Show(keypos)

                ' End If



                ' MessageBox.Show(row.Cells(0).Value)

                ' insertcommand.Parameters.Clear()
                ' insertcommand.CommandText = query
                ''  insertcommand.Parameters.AddWithValue("@TaskID", row.Cells(0).Value)
                '  insertcommand.Parameters.AddWithValue("@Complete", "False")
                '  insertcommand.Parameters.AddWithValue("@Task", row.Cells(1).Value)
                ' insertcommand.Parameters.AddWithValue("@Start_date", row.Cells(2).Value)
                '  insertcommand.Parameters.AddWithValue("@Due_Date", row.Cells(3).Value)
                ' insertcommand.Parameters.AddWithValue("@JRID", txtJRID.Text)
                ' insertcommand.Parameters.AddWithValue("@Task_Manager", row.Cells(4).Value)
                ''  insertcommand.Parameters.AddWithValue("@Entered_By", GetUserName())
                '  insertcommand.Parameters.AddWithValue("@Time_Entered", now)
                ' insertcommand.ExecuteNonQuery()
            End If
            keypos = keypos + 1
        Next
    End Sub

    Function GetSettingsCRM()

        bchk = False

        Dim fSettingsXml As String() = Directory.GetFiles(getCRMFolderXml, "a.xml")

        For Each FileName In fSettingsXml

            If (System.IO.File.Exists(FileName)) Then
                'Read Settings 
                If FileName = getCRMFolderXml & "" & "a.xml" Then
                    ' Get Filename and TraID
                    ValidateXMLCRMSettings(FileName)

                    ' Check if the file is valid 
                    If (getOName <> Orig_FileName) Then

                        msgstatus = "File Is invalid Or renamed, server checked at " & getServerTime()

                        'If Not boolDc Then
                        'SQL.LiveExecQuery("Update files Set uploaded = 2, status = '" & msgstatus & "'  where file_pk = " & getFId)
                        'If NotEmpty(SQL.Exception) Then MsgBox(SQL.Exception) : Exit Sub
                        'Else
                        'Call deleteExFiles(getOName)
                        'End If

                        Call deleteExFiles(getFolderZip)

                        Call setEmailSend(getOName, "", getEmail, "", "", "Sales Consolidator", msgstatus, 0)
                        getOName = Nothing

                        bchk = True
                        'rst10sec()
                        'Exit Sub

                    End If

                    ' Check if the file was already uploaded
                    If GetFilename(Orig_FileName) Then

                        If Not boolDc Then
                            msgstatus = "File Name already uploaded, server checked at " & getServerTime()
                            'SQL.LiveExecQuery("Update files set uploaded = 3, status = '" & msgstatus & "'  where file_pk = " & getFId)
                            'If NotEmpty(SQL.Exception) Then MsgBox(SQL.Exception) : Exit Sub
                        Else
                            msgstatus = "File Name already transmitted, server checked at " & getServerTime()
                            ' Call deleteFtpZipFile(getFolderZip)
                            Call deleteExFiles(getFolderZip)
                        End If

                        Call setEmailSend(getOName, "", getEmail, "", "", "Sales Consolidator", msgstatus, 0)


                        getOName = Nothing
                        bchk = True

                        'rst10sec()
                        'Exit Sub
                    End If

                    'ReadXmlSettingsInv(FileName)

                End If

            End If

        Next

        Return bchk

    End Function


    Public Sub getPayments(csvPayments As String)

        Dim TextFileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(csvPayments)

        TextFileReader.TextFieldType = FileIO.FieldType.Delimited
        TextFileReader.SetDelimiters(",")

        Dim TextFileTable As DataTable = Nothing

        Dim Column As DataColumn
        Dim Row As DataRow
        Dim UpperBound As Int32
        Dim ColumnCount As Int32
        Dim CurrentRow As String()

        While Not TextFileReader.EndOfData
            Try
                CurrentRow = TextFileReader.ReadFields()

                If TextFileTable Is Nothing Then
                    TextFileTable = New DataTable("TextFileTable")
                    ''# Get number of columns
                    UpperBound = CurrentRow.GetUpperBound(0)
                    ''# Create new DataTable
                    For ColumnCount = 0 To 7
                        Column = New DataColumn()
                        Column.DataType = System.Type.GetType("System.String")
                        Column.ColumnName = "Column" & ColumnCount
                        Column.Caption = "Column" & ColumnCount
                        Column.ReadOnly = True
                        Column.Unique = False
                        TextFileTable.Columns.Add(Column)
                    Next
                End If


                If Not CurrentRow Is Nothing Then
                    ''# Check if DataTable has been created

                    Row = TextFileTable.NewRow
                    For ColumnCount = 0 To 7

                        Try
                            Row("Column" & ColumnCount) = CurrentRow(ColumnCount).ToString
                        Catch

                        End Try

                    Next
                    TextFileTable.Rows.Add(Row)
                End If
            Catch ex As _
            Microsoft.VisualBasic.FileIO.MalformedLineException
                MsgBox("Line " & ex.Message &
                "is not valid and will be skipped.")
            End Try
        End While
        TextFileReader.Dispose()
        'frmMain.DataGrid1.DataSource = TextFileTable
        dgvpayments.DataSource = TextFileTable


    End Sub



    Public Sub getTransaction(csvTrans As String)

        Dim TextFileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(csvTrans)

        TextFileReader.TextFieldType = FileIO.FieldType.Delimited
        TextFileReader.SetDelimiters(",")

        Dim TextFileTable As DataTable = Nothing

        Dim Column As DataColumn
        Dim Row As DataRow
        Dim UpperBound As Int32
        Dim ColumnCount As Int32
        Dim CurrentRow As String()

        While Not TextFileReader.EndOfData
            Try
                CurrentRow = TextFileReader.ReadFields()

                If TextFileTable Is Nothing Then
                    TextFileTable = New DataTable("TextFileTable")
                    ''# Get number of columns
                    UpperBound = CurrentRow.GetUpperBound(0)
                    ''# Create new DataTable
                    For ColumnCount = 0 To 7
                        Column = New DataColumn()
                        Column.DataType = System.Type.GetType("System.String")
                        Column.ColumnName = "Column" & ColumnCount
                        Column.Caption = "Column" & ColumnCount
                        Column.ReadOnly = True
                        Column.Unique = False
                        TextFileTable.Columns.Add(Column)
                    Next
                End If


                If Not CurrentRow Is Nothing Then
                    ''# Check if DataTable has been created

                    Row = TextFileTable.NewRow
                    For ColumnCount = 0 To 7

                        Try
                            Row("Column" & ColumnCount) = CurrentRow(ColumnCount).ToString
                        Catch

                        End Try

                    Next
                    TextFileTable.Rows.Add(Row)
                End If
            Catch ex As _
            Microsoft.VisualBasic.FileIO.MalformedLineException
                MsgBox("Line " & ex.Message &
                "is not valid and will be skipped.")
            End Try
        End While
        TextFileReader.Dispose()
        'frmMain.DataGrid1.DataSource = TextFileTable
        dgvTrans.DataSource = TextFileTable



    End Sub
    Public Sub getFinishedSales(csvSales As String)

        Dim TextFileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(csvSales)

        TextFileReader.TextFieldType = FileIO.FieldType.Delimited
        TextFileReader.SetDelimiters(",")

        Dim TextFileTable As DataTable = Nothing

        Dim Column As DataColumn
        Dim Row As DataRow
        Dim UpperBound As Int32
        Dim ColumnCount As Int32
        Dim CurrentRow As String()

        While Not TextFileReader.EndOfData
            Try
                CurrentRow = TextFileReader.ReadFields()

                If TextFileTable Is Nothing Then
                    TextFileTable = New DataTable("TextFileTable")
                    ''# Get number of columns
                    UpperBound = CurrentRow.GetUpperBound(0)
                    ''# Create new DataTable
                    For ColumnCount = 0 To 7
                        Column = New DataColumn()
                        Column.DataType = System.Type.GetType("System.String")
                        Column.ColumnName = "Column" & ColumnCount
                        Column.Caption = "Column" & ColumnCount
                        Column.ReadOnly = True
                        Column.Unique = False
                        TextFileTable.Columns.Add(Column)
                    Next
                End If


                If Not CurrentRow Is Nothing Then
                    ''# Check if DataTable has been created

                    Row = TextFileTable.NewRow
                    For ColumnCount = 0 To 7

                        Try
                            Row("Column" & ColumnCount) = CurrentRow(ColumnCount).ToString
                        Catch

                        End Try

                    Next
                    TextFileTable.Rows.Add(Row)
                End If
            Catch ex As _
            Microsoft.VisualBasic.FileIO.MalformedLineException
                MsgBox("Line " & ex.Message &
                "is not valid and will be skipped.")
            End Try
        End While
        TextFileReader.Dispose()
        'frmMain.DataGrid1.DataSource = TextFileTable
        dgvSales.DataSource = TextFileTable



    End Sub
    Public Sub ValidateXMLCRMSettings(xmlFileName As String)

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlFileName)
        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/Settings")
        For Each node As XmlNode In nodes

            Orig_FileName = node.SelectSingleNode("FileName").InnerText

            StoreName = node.SelectSingleNode("StoreName").InnerText

            StoreCode = node.SelectSingleNode("StoreCode").InnerText

            getEmail = node.SelectSingleNode("Email").InnerText

            getUploadStamp = node.SelectSingleNode("DateTime").InnerText

            getEmpname = node.SelectSingleNode("User").InnerText

            DateTran = Convert.ToDateTime(node.SelectSingleNode("DateTran").InnerText)

            DTranDay = node.SelectSingleNode("DTranDay").InnerText

            'getCSVFininsedSales = getCRMFolderXml + "Z004_" + Convert.ToString(Convert.ToDateTime(DateTran).Today.Day)
            '‪C:\GetCRMExported\Z004_26.CSV

            getCSVFininsedSales = getCRMFolderXml + "Z004_" + DTranDay + ".CSV"
            getCSVTrans = getCRMFolderXml + "Z002_" + DTranDay + ".CSV"
            getCSVPayments = getCRMFolderXml + "Z001_" + DTranDay + ".CSV"

        Next

    End Sub

    Public Sub deleteFtpZipFileCRMSales(Orig_FileName As String)
        Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://ftp.cindysbakery.com/" + Orig_FileName), System.Net.FtpWebRequest)
        request.Credentials = New System.Net.NetworkCredential("crm@cindysbakery.com", "csd@cindys")
        request.Method = System.Net.WebRequestMethods.Ftp.DeleteFile
        request.GetResponse()
    End Sub

    'Email Function
    Private Sub setEmailSend(sSubject As String, sBody As String, sTo As String, sCC As String, sFilename As String, sDisplayname As String, sTatus As String, sType As Int32)
        Dim oApp As Interop.Outlook._Application
        oApp = New Interop.Outlook.Application

        Dim oMsg As Interop.Outlook._MailItem
        oMsg = oApp.CreateItem(Interop.Outlook.OlItemType.olMailItem)

        Dim usrstore, userup, datestr, ticketstr, remarkstr, dateup, spacex, spacey, endstr, htmlend As String

        'Create Messages
        datestr = "<html><body><pre wrap>Date/Time: " & now

        'Create Messages
        datestr = "<html><body><pre wrap>Date/Time: " & now

        If sType = 2 Then
            'dateup = "<br/>Sales Data Uploaded:  <b>" & " " & getUploadStamp & " " & "</b>"
            ' ElseIf sType = 2 Then
            dateup = "<br/>CRM Sales Data Transmitted:  <b>" & " " & getUploadStamp & " " & "</b>"
            ' ElseIf sType = 3 Then
            'dateup = "<br/>Inventory Data Uploaded:  <b>" & " " & getUploadStamp & " " & "</b>"
            ' ElseIf sType = 4 Then
            'dateup = "<br/>Inventory Data Transmitted:  <b>" & " " & getUploadStamp & " " & "</b>"
        ElseIf sType = 3 Then
            dateup = "<br/>Process:  <b>" & " " & "File is currupted please re-transmit again." & " " & "</b>"
        Else
            dateup = "<br/>Process:  <b>" & " " & getUploadStamp & " " & "</b>"
        End If

        userup = "<br/>User:  <b>" & " " & getEmpname & " " & "</b>"
        usrstore = "<br/>Store:  <b>" & " " & StoreCode + " - " + StoreName & " " & "</b>"
        ticketstr = "<br/>File: <b>" & " " & sSubject & " " & "</b>"
        remarkstr = "<br/>Remarks: <b>" & " " & sTatus & " " & "</b>"
        spacex = "<br/>"
        spacey = "<br/>"
        endstr = "<br/>*** This is an automatically generated email, please do not reply ***"
        htmlend = "</pre></body></html>"


        oMsg.Subject = sSubject
        oMsg.HTMLBody = datestr + dateup + userup + ticketstr + remarkstr + spacex + spacey + endstr + htmlend

        oMsg.To = sTo
        'oMsg.CC = "j.tabong@cindysbakery.com"

        Dim strS As String = sFilename
        Dim strN As String = sDisplayname
        If sFilename <> "" Then
            Dim sBodyLen As Integer = Int(sBody.Length)
            Dim oAttachs As Interop.Outlook.Attachments = oMsg.Attachments
            Dim oAttach As Interop.Outlook.Attachment

            oAttach = oAttachs.Add(strS, , sBodyLen, strN)

        End If

        oMsg.Send()

    End Sub

    Public Sub ExFile()
        Try
            Using zip As ZipFile = ZipFile.Read(getFolderZip + getOName)
                For Each zip_entry As ZipEntry In zip
                    zip_entry.Password = "johnnyuser2"
                    zip_entry.Extract(getCRMFolderXml)
                Next zip_entry
            End Using

        Catch ex As Exception


            'MessageBox.Show("Error extracting archive.\n" + ex.Message) : Exit Sub
            'Exit Sub

            Call setEmailSend(getOName, "", "consolidator@cindysbakery.com", "", "", "CRM Consolidator", "File Currupted", 3)

            Call deleteFtpZipFile(getOName)

            Application.Exit()

        End Try
    End Sub

    'Delete Zip/Xml Files
    Public Sub deleteExFiles(getCRMFolderXml As String)
        'Delete created xml sales file
        'For Each deleteFile In Directory.GetFiles(getCRMFolderXml, "*.xml", SearchOption.TopDirectoryOnly)
        'File.Delete(deleteFile)
        'Next

        'Delete zip file
        For Each deleteFile In Directory.GetFiles(getCRMFolderXml, getOName, SearchOption.TopDirectoryOnly)
            File.Delete(deleteFile)
        Next

        'Delete created xml inventory file
        'For Each deleteFile In Directory.GetFiles(getCRMFolderXml, "*I.xml", SearchOption.TopDirectoryOnly)
        'File.Delete(deleteFile)
        'Next

        'Delete zip file
        'For Each deleteFile In Directory.GetFiles(getCRMFolderXml, "*.csv", SearchOption.TopDirectoryOnly)
        'File.Delete(deleteFile)
        'Next

    End Sub


    Public Sub deleteExFilesX(getCRMFolderXml As String)
        'Delete created xml sales file
        For Each deleteFile In Directory.GetFiles(getCRMFolderXml, "*.xml", SearchOption.TopDirectoryOnly)
            File.Delete(deleteFile)
        Next

        'Delete zip file
        For Each deleteFile In Directory.GetFiles(getCRMFolderXml, "*.zip", SearchOption.TopDirectoryOnly)
            File.Delete(deleteFile)
        Next

        'Delete zip file
        For Each deleteFile In Directory.GetFiles(getCRMFolderXml, "*.csv", SearchOption.TopDirectoryOnly)
            File.Delete(deleteFile)
        Next

    End Sub

    Public Function GetMDB()

        Try
            SQL.DBCon3.Open()
            SQL.DBCon3.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    'Set Timer to 5 mins

End Class