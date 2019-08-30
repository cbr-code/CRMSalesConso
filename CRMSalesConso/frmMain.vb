Public Class frmMain
    Dim SQL As New SQLControl
    Dim SQL2 As New SQLControl
    Private Sub ExtractSalesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'frmExtract.ShowDialog()
    End Sub

    Private Sub ServerSetupToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Module Underconstruction.", "Store Backend", MessageBoxButtons.OK)
    End Sub

    Private Sub GetUpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetUpdateToolStripMenuItem.Click
        ' MessageBox.Show("Module Underconstruction.", "Store Backend", MessageBoxButtons.OK)
        'GetZip.ShowDialog()
        'FrmLogs.ShowDialog()
    End Sub

    Private Sub StoreInfoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'frmStore.ShowDialog()
    End Sub

    Private Sub UserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserToolStripMenuItem.Click
        'frmUpdateUser.ShowDialog()

        'Dim n As String = MsgBox("Update Users Online", MsgBoxStyle.YesNo, "Store Backend")
        'If n = vbYes Then


        '    ' EXECUTE QUERY
        '    SQL2.ExecQuery("SELECT EmpTab.EmpNo, EmpTab.empname , pos.positionname FROM EmpTab left join PositionTab pos on pos.positionid = EmpTab.positionid where pos.positionid in (85, 48, 278, 18, 319, 300, 205, 206, 225, 191, 259)")


        '    ' REPORT & ABORT ON ERRORS
        '    If NotEmpty(SQL2.Exception) Then MsgBox(SQL2.Exception) : Exit Sub


        '    For Each r As DataRow In SQL2.DBDT.Rows
        '        'lvTables.Items.Add(r("TABLE_NAME"))



        '        SQL.LiveExecQuery("SELECT * from users where ucode = '" & r("EmpNo") & "' ")

        '        If SQL.RecordCount = 0 Then

        '            Dim ucode As String = r("EmpNo")
        '            Dim uname As String = r("empname")
        '            Dim upos As String = r("positionname")


        '            SQL.LiveExecQuery("INSERT INTO users (ucode,uname,upos) values ('" & ucode & "', '" & uname & "', '" & upos & "')")



        '            If NotEmpty(SQL.Exception) Then MsgBox(SQL.Exception) : Exit Sub

        '        End If


        '    Next

        'End If


        'MessageBox.Show("Succesfull update.", "Store Backend", MessageBoxButtons.OK)


    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' If uid = 0 Then
        'Dim frmLogin As New frmLogin
        'mstrip.Hide()
        'ToolStrip1.Hide()
        ' frmLogin.ShowDialog()
        ' End If

        'GetZip.Timer1.Start()

        'GetData.Timer1.Start()

        'ProdCode.ShowDialog()
        'cscParser.ShowDialog()

        'cscParser.Timer1.Start()

        'inv.Timer1.Start()


    End Sub

    Private Sub LogOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogOutToolStripMenuItem.Click
        Dim n As String = MsgBox("Quit Application", MsgBoxStyle.YesNo, "Store Backend")
        If n = vbYes Then
            'Application.Exit()
        End If
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click

    End Sub

    Private Sub StartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartToolStripMenuItem.Click

        'GetZip.ShowDialog()

    End Sub

    Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize

        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            NotifyIcon1.Visible = True
            'NotifyIcon1.ShowBalloonTip(5000)
            'NotifyIcon1.BalloonTipText = "POS Consolidator working on background"
            'NotifyIcon1.BalloonTipTitle = "POS HQ Consolidator"
            NotifyIcon1.ShowBalloonTip(8000, "Consolidator", "Consolidator working on background", ToolTipIcon.Info)
        End If

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Maximized
        Me.NotifyIcon1.Visible = False
    End Sub
End Class