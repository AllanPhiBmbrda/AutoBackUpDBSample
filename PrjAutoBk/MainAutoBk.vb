Imports System.IO
Imports System.Threading
Imports System.ComponentModel
Imports System.Globalization
Imports MySql.Data.MySqlClient
Imports System.IO.Compression
Imports Ionic.Zip
Imports System.Net.Mail


Public Class MainAutoBk

    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams

        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get

    End Property


    Dim AllocateZip As String = Nothing
    Dim SaveZip As String = Nothing
    Dim AttachmentForMail As String
    Private Sub MainAutoBk_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        InitialFolderLook()
        AuBkGrid01Header()
        AuBkGrid02Header()
        'LostFileTextCall()
        'LostLogFile()
        ReadTextFileMode()
        ReadLogFileMode()

        AuBkDP01.Value = DateTime.ParseExact("10:00:00 AM", "hh:mm:ss tt", CultureInfo.InvariantCulture)
        AuBkDP02.Value = DateTime.ParseExact("04:00:00 PM", "hh:mm:ss tt", CultureInfo.InvariantCulture)

        AntiPeeping()

        Me.Size = New Drawing.Size(410, 150)
        TabControl1.TabPages(0).Enabled = False
        TabControl1.TabPages(1).Enabled = False
        TabControl1.TabPages(2).Enabled = False
        TabControl1.SelectTab(3)

    End Sub

    Sub AntiPeeping()

        Select Case GetKeyNow

            Case True

                For a = 0 To AuBkGrid01.Rows.Count - 1
                    AuBkGrid01.Columns(1).DefaultCellStyle.BackColor = Color.White
                    AuBkGrid01.Columns(1).DefaultCellStyle.ForeColor = Color.Black
                Next

            Case False

                For a = 0 To AuBkGrid01.Rows.Count - 1
                    AuBkGrid01.Columns(1).DefaultCellStyle.BackColor = Color.White
                    AuBkGrid01.Columns(1).DefaultCellStyle.ForeColor = Color.White
                Next

        End Select
        AuBkGrid01.ClearSelection()

    End Sub
    Sub InitialFolderLook()
        If (Not System.IO.Directory.Exists(Application.StartupPath + "\Database List")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath + "\Database List")
        End If
        If (Not System.IO.Directory.Exists(Application.StartupPath + "\Connection Tru Text")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath + "\Connection Tru Text")
        End If
    End Sub
    Sub AuBkGrid01Header()
        With AuBkGrid01

            .Rows.Clear()
            .Columns.Clear()
            .Columns.Add("col00", "Object Name")
            .Columns.Add("col01", "Connection String")
            .Columns.Add("col02", "Remark")
            .Columns(0).Width = 150
            .Columns(1).Width = 350
            .Columns(2).Width = 250

        End With
    End Sub

    Sub AuBkGrid02Header()
        With AuBkGrid02
            .Rows.Clear()
            .Columns.Clear()
            .Columns.Add("col00", "Time")
            .Columns.Add("col01", "Object Name")
            .Columns.Add("col02", "LOG CODE / LOG MESSAGE")
            .Columns(0).Width = 150
            .Columns(1).Width = 250
            .Columns(2).Width = 350

        End With

    End Sub

    Private Sub AuBkBtn03_Click(sender As Object, e As EventArgs) Handles AuBkBtn03.Click
        Select Case AuBkBtn03.Text
            Case "ADD"
                AuBk01.Enabled = True
                AuBk02.Enabled = True
                AuBk03.Enabled = True
                AuBk04.Enabled = True
                AuBk05.Enabled = True
                AuBk06.Enabled = True
                AuBkBtn03.Text = "LOCK"
            Case "LOCK"
                AuBk01.Enabled = False
                AuBk02.Enabled = False
                AuBk03.Enabled = False
                AuBk04.Enabled = False
                AuBk05.Enabled = False
                AuBk06.Enabled = False
                AuBkBtn03.Text = "ADD"
        End Select

    End Sub

    Private Sub AuBkBtn04_Click(sender As Object, e As EventArgs) Handles AuBkBtn04.Click
        If AuBkBtn03.Text = "LOCK" Then
            If AuBk01.Text = Nothing Then
                MessageBox.Show("Insert Name of Object")
            ElseIf AuBk02.Text = Nothing Then
                MessageBox.Show("Insert Name of SERVER")
            ElseIf AuBk03.Text = Nothing Then
                MessageBox.Show("Insert PORT NUMBER")
            ElseIf AuBk04.Text = Nothing Then
                MessageBox.Show("Insert Name of USER/UID")
            ElseIf AuBk05.Text = Nothing Then
                MessageBox.Show("Insert USERNAME")
            ElseIf AuBk06.Text = Nothing Then
                MessageBox.Show("Insert PASSWORD")
            ElseIf AuBk06.Text = Nothing Then
                MessageBox.Show("Insert DATABASE")
            Else
                AuBkGrid01.Rows.Add(AuBk01.Text, "server=" & AuBk03.Text & ";port=" & AuBk04.Text & ";user=" & AuBk05.Text & ";pwd=" & AuBk06.Text & ";database=" & AuBk02.Text)
            End If
        End If
        AuBkBtn03.PerformClick()
    End Sub

    Dim LinkTxtFile As String
    Dim LogTxtFile As String
    Sub LostFileTextCall()
        LinkTxtFile = Application.StartupPath + "\Connection Tru Text\ConnectionStrings.txt"
        If Not File.Exists(LinkTxtFile) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(LinkTxtFile)

            End Using
        End If
    End Sub

    Sub LostLogFile()
        LogTxtFile = Application.StartupPath + "\Connection Tru Text\LogString.txt"
        If Not File.Exists(LogTxtFile) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(LogTxtFile)

            End Using
        End If
    End Sub

    Sub SaveConnectListCall()

        Dim SaveLinkTxtFile As String
        SaveLinkTxtFile = Application.StartupPath + "\Connection Tru Text\ConnectionStrings.txt"

        If File.Exists(SaveLinkTxtFile) Then
            File.Delete(SaveLinkTxtFile)
        End If

        If Not File.Exists(SaveLinkTxtFile) Then
            Using sw As StreamWriter = File.CreateText(SaveLinkTxtFile)
                For i = 0 To AuBkGrid01.Rows.Count - 1
                    sw.WriteLine(Encrypt(AuBkGrid01(0, i).Value.ToString & "#" & AuBkGrid01(1, i).Value.ToString))
                Next
            End Using
        End If

    End Sub

    Sub SaveLogListCall()

        Dim SaveLogTxtFile As String
        SaveLogTxtFile = Application.StartupPath + "\Connection Tru Text\LogString.txt"

        If File.Exists(SaveLogTxtFile) Then
            File.Delete(SaveLogTxtFile)
        End If

        If Not File.Exists(SaveLogTxtFile) Then
            Using sw As StreamWriter = File.CreateText(SaveLogTxtFile)
                For i = 0 To AuBkGrid02.Rows.Count - 1
                    sw.WriteLine(Encrypt(AuBkGrid02(0, i).Value.ToString & "#" & AuBkGrid02(1, i).Value.ToString & "#" & AuBkGrid02(2, i).Value.ToString))
                Next
            End Using
        End If

    End Sub

    Dim LoadLinkName As String
    Dim TextLineRead As String = Nothing
    Dim SplitLine() As String
    Dim SplitLine2 As String

    Dim LoadLogName As String

    Sub ReadTextFileMode()

        LoadLinkName = Application.StartupPath + "\Connection Tru Text\ConnectionStrings.txt"
        If System.IO.File.Exists(LoadLinkName) = True Then

            Dim objReader As New System.IO.StreamReader(LoadLinkName)

            Do While objReader.Peek() <> -1
                TextLineRead = objReader.ReadLine()
                SplitLine = Split(Decrypt(TextLineRead), "#")
                AuBkGrid01.Rows.Add(SplitLine)
            Loop

            objReader.Close()
            objReader = Nothing

        End If

    End Sub

    Sub ReadLogFileMode()

        LoadLogName = Application.StartupPath + "\Connection Tru Text\LogString.txt"
        If System.IO.File.Exists(LoadLogName) = True Then

            Dim objReader As New System.IO.StreamReader(LoadLogName)

            Do While objReader.Peek() <> -1
                TextLineRead = objReader.ReadLine()
                SplitLine = Split(Decrypt(TextLineRead), "#")
                AuBkGrid02.Rows.Add(SplitLine)
            Loop

            objReader.Close()
            objReader = Nothing

        End If

    End Sub

    Private Sub AuBkBtn02_Click(sender As Object, e As EventArgs) Handles AuBkBtn02.Click
        SaveConnectListCall()
        Try
            AuBkGrid01.Rows.Clear()
            ReadTextFileMode()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Sub ExportModeA()
        If (Not System.IO.Directory.Exists(Application.StartupPath + "\Database List\Database A " + Today.ToString("ddMMMyyyy"))) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath + "\Database List\Database A " + Today.ToString("ddMMMyyyy"))
        End If

        Dim ActCon(100) As Boolean ' Only 100 Connection String only

        For b = 0 To AuBkGrid01.Rows.Count - 1

            Try
                Dim ToTestConn As MySqlConnection = New MySqlConnection(AuBkGrid01(1, b).Value.ToString)
                ToTestConn.Open()
                ActCon(b) = True
                ToTestConn.Close()

            Catch ex As Exception
                ActCon(b) = False
            End Try

        Next

        For i = 0 To AuBkGrid01.Rows.Count - 1
            Try
                Select Case ActCon(i)
                    Case True
                        Dim SaveLink As String = Application.StartupPath + "\Database List\Database A " + Today.ToString("ddMMMyyyy") + "\" + AuBkGrid01(0, i).Value.ToString + " A " + Today.ToString("ddMMMyyyy") + ".sql"
                        Dim ToExportConn As MySqlConnection = New MySqlConnection(AuBkGrid01(1, i).Value.ToString)
                        Dim ToExportCmd As MySqlCommand = New MySqlCommand
                        ToExportCmd.Connection = ToExportConn
                        ToExportConn.Open()
                        Dim ToExportMb As MySqlBackup = New MySqlBackup(ToExportCmd)
                        ToExportMb.ExportToFile(SaveLink)
                        ToExportConn.Close()

                        AuBkGrid01.Invoke(DirectCast(Sub() AuBkGrid01(2, i).Value = "Export [DONE]", MethodInvoker))
                        AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), AuBkGrid01(0, i).Value.ToString, "Export [DONE]"), MethodInvoker))
                    Case False
                        AuBkGrid01.Invoke(DirectCast(Sub() AuBkGrid01(2, i).Value = "Your REQUEST is Not connected to Corresponding SERVER", MethodInvoker))
                        AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), AuBkGrid01(0, i).Value.ToString, "Your REQUEST is Not connected to Corresponding SERVER"), MethodInvoker))
                End Select

            Catch ex As Exception
                AuBkGrid01.Invoke(DirectCast(Sub() AuBkGrid01(2, i).Value = ex.Message, MethodInvoker))
            End Try
        Next

        Select Case AuBkChk01.Checked
            Case True
                Try
                    Using GetMyZip As New ZipFile
                        With GetMyZip

                            Dim GetDirSQL As New DirectoryInfo(Application.StartupPath + "\Database List\Database A " & Today.ToString("ddMMMyyyy"))
                            Dim FileMakeArray As FileInfo() = GetDirSQL.GetFiles("*.sql")
                            Dim LoopFileMode As FileInfo
                            For Each LoopFileMode In FileMakeArray
                                .AddFile(LoopFileMode.FullName, "")
                                .Password = "universalgloves2016"
                            Next LoopFileMode

                            If File.Exists(Application.StartupPath + "\Database List" + "\ZIP A " + Today.ToString("ddMMMyyyy") + ".zip") Then
                                File.Delete(Application.StartupPath + "\Database List" + "\ZIP A " + Today.ToString("ddMMMyyyy") + ".zip")
                            End If
                            .Save(Application.StartupPath + "\Database List" + "\ZIP A " + Today.ToString("ddMMMyyyy") + ".zip")
                            AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Get to ZIP A", "SUCCESS"), MethodInvoker))
                        End With
                    End Using
                Catch ex As Exception
                    AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString, "Get to ZIP A", "Compression of FILE : FAILED"), MethodInvoker))
                End Try
                If (System.IO.Directory.Exists(Application.StartupPath + "\Database List\Database A " + Today.ToString("ddMMMyyyy"))) Then
                    System.IO.Directory.Delete(Application.StartupPath + "\Database List\Database A " + Today.ToString("ddMMMyyyy"), True)
                End If

                Select Case AuBkChk02.Checked

                    Case True
                        Try
                            AttachmentForMail = Application.StartupPath + "\Database List" + "\ZIP A " + Today.ToString("ddMMMyyyy") + ".zip"
                            SendtoEmail()
                        Catch ex As Exception
                            AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Sending ZIP A to EMAIL", "STATUS : FAILED"), MethodInvoker))
                        End Try

                End Select

                Try
                    Select Case AuBkChk03.Checked
                        Case True
                            If System.IO.Directory.Exists("\\192.168.2.200\server ug\KC PC") Then
                                File.Copy(Application.StartupPath + "\Database List" + "\ZIP A " + Today.ToString("ddMMMyyyy") + ".zip", "\\192.168.2.200\server ug\KC PC" & "\ZIP A " + Today.ToString("ddMMMyyyy") + ".zip")
                            Else
                                AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Directory to File Server", "Saving Failed"), MethodInvoker))
                            End If
                    End Select
                Catch ex As Exception
                End Try

        End Select
        SaveLogListCall()
        RefreshingMode()
    End Sub

    Sub ExportModeB()

        If (Not System.IO.Directory.Exists(Application.StartupPath + "\Database List\Database B " + Today.ToString("ddMMMyyyy"))) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath + "\Database List\Database B " + Today.ToString("ddMMMyyyy"))
        End If

        Dim ActCon(100) As Boolean ' Only 100 Connection String only

        For b = 0 To AuBkGrid01.Rows.Count - 1

            Try
                Dim ToTestConn As MySqlConnection = New MySqlConnection(AuBkGrid01(1, b).Value.ToString)
                ToTestConn.Open()
                ActCon(b) = True
                ToTestConn.Close()

            Catch ex As Exception
                ActCon(b) = False

            End Try
        Next

        For i = 0 To AuBkGrid01.Rows.Count - 1
            Try
                Select Case ActCon(i)
                    Case True
                        Dim SaveLink As String = Application.StartupPath + "\Database List\Database B " + Today.ToString("ddMMMyyyy") + "\" + AuBkGrid01(0, i).Value.ToString + " B " + Today.ToString("ddMMMyyyy") + ".sql"
                        Dim ToExportConn As MySqlConnection = New MySqlConnection(AuBkGrid01(1, i).Value.ToString)
                        Dim ToExportCmd As MySqlCommand = New MySqlCommand
                        ToExportCmd.Connection = ToExportConn
                        ToExportConn.Open()
                        Dim ToExportMb As MySqlBackup = New MySqlBackup(ToExportCmd)
                        ToExportMb.ExportToFile(SaveLink)
                        ToExportConn.Close()
                        AuBkGrid01.Invoke(DirectCast(Sub() AuBkGrid01(2, i).Value = "Export [DONE]", MethodInvoker))
                        AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), AuBkGrid01(0, i).Value.ToString, "Export [DONE]"), MethodInvoker))
                    Case False
                        AuBkGrid01.Invoke(DirectCast(Sub() AuBkGrid01(2, i).Value = "Your REQUEST is Not connected to Corresponding SERVER", MethodInvoker))
                        AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), AuBkGrid01(0, i).Value.ToString, "Your REQUEST is Not connected to Corresponding SERVER"), MethodInvoker))
                End Select

            Catch ex As Exception
                AuBkGrid01.Invoke(DirectCast(Sub() AuBkGrid01(2, i).Value = ex.Message, MethodInvoker))
            End Try
        Next

        Select Case AuBkChk01.Checked
            Case True
                Try
                    Using GetMyZip As New ZipFile
                        With GetMyZip

                            Dim GetDirSQL As New DirectoryInfo(Application.StartupPath + "\Database List\Database B " & Today.ToString("ddMMMyyyy"))
                            Dim FileMakeArray As FileInfo() = GetDirSQL.GetFiles("*.sql")
                            Dim LoopFileMode As FileInfo
                            For Each LoopFileMode In FileMakeArray
                                .AddFile(LoopFileMode.FullName, "")
                                .Password = "universalgloves2016"
                            Next LoopFileMode

                            If File.Exists(Application.StartupPath + "\Database List" + "\ZIP B " + Today.ToString("ddMMMyyyy") + ".zip") Then
                                File.Delete(Application.StartupPath + "\Database List" + "\ZIP B " + Today.ToString("ddMMMyyyy") + ".zip")
                            End If
                            .Save(Application.StartupPath + "\Database List" + "\ZIP B " + Today.ToString("ddMMMyyyy") + ".zip")
                            AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Get to ZIP B", "SUCCESS"), MethodInvoker))
                        End With
                    End Using
                Catch ex As Exception
                    AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Get to ZIP B", "Compression of FILE : FAILED"), MethodInvoker))
                End Try
                If (System.IO.Directory.Exists(Application.StartupPath + "\Database List\Database B " + Today.ToString("ddMMMyyyy"))) Then
                    System.IO.Directory.Delete(Application.StartupPath + "\Database List\Database B " + Today.ToString("ddMMMyyyy"), True)
                End If

                Select Case AuBkChk02.Checked

                    Case True
                        Try
                            AttachmentForMail = Application.StartupPath + "\Database List" + "\ZIP B " + Today.ToString("ddMMMyyyy") + ".zip"
                            SendtoEmail()
                        Catch ex As Exception
                            AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Sending ZIP B to EMAIL", "STATUS : FAILED"), MethodInvoker))
                        End Try

                End Select
                Try
                    Select Case AuBkChk03.Checked
                        Case True
                            If System.IO.Directory.Exists("\\192.168.2.200\server ug\KC PC") Then
                                File.Copy(Application.StartupPath + "\Database List" + "\ZIP B " + Today.ToString("ddMMMyyyy") + ".zip", "\\192.168.2.200\server ug\KC PC" & "\ZIP B " + Today.ToString("ddMMMyyyy") + ".zip")
                            Else
                                AuBkGrid02.Invoke(DirectCast(Sub() AuBkGrid02.Rows.Add(Today.ToString("dd MMM yyyy"), "Directory to File Server", "Saving Failed"), MethodInvoker))
                            End If
                    End Select
                Catch ex As Exception
                End Try

        End Select
        SaveLogListCall()
        RefreshingMode()

    End Sub

    Private Sub AuBkBtn01_Click(sender As Object, e As EventArgs) Handles AuBkBtn01.Click
        AuBkGrid01Header()
        ReadTextFileMode()
        AntiPeeping()
    End Sub

    Sub RefreshingMode()
        AuBkGrid01Header()
        ReadTextFileMode()
        AntiPeeping()
    End Sub
    Private Sub AuBkBtn05_Click(sender As Object, e As EventArgs) Handles AuBkBtn05.Click
        AuBkBtn05.Enabled = False
        ExportModeA()
        'OnClickTheWorker2()
    End Sub

#Region "BGW for EXPORTING"
    Private BGWorkMode() As BackgroundWorker
    Private i = 0
    Private g = 0

    Sub OnClickTheWorker()

        i += 1
        ReDim BGWorkMode(i)
        BGWorkMode(i) = New BackgroundWorker
        BGWorkMode(i).WorkerReportsProgress = True
        BGWorkMode(i).WorkerSupportsCancellation = True
        AddHandler BGWorkMode(i).DoWork, AddressOf WorkerDoWork
        AddHandler BGWorkMode(i).ProgressChanged, AddressOf WorkerProgressChanged
        AddHandler BGWorkMode(i).RunWorkerCompleted, AddressOf WorkerCompleted
        BGWorkMode(i).RunWorkerAsync()
        If Not BGWorkMode(i).IsBusy Then
            BGWorkMode(i).Dispose()
            i = 0
        End If

    End Sub
    Private Sub WorkerDoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        AuBkTimer01.Enabled = False
        ExportModeB()

    End Sub
    Private Sub WorkerProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub
    Private Sub WorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        'AuBkBtn05.Enabled = True
        AuBkTimer01.Enabled = True

    End Sub

    Sub OnClickTheWorker2()

        g += 1
        ReDim BGWorkMode(g)
        BGWorkMode(g) = New BackgroundWorker
        BGWorkMode(g).WorkerReportsProgress = True
        BGWorkMode(g).WorkerSupportsCancellation = True
        AddHandler BGWorkMode(g).DoWork, AddressOf WorkerDoWork2
        AddHandler BGWorkMode(g).ProgressChanged, AddressOf WorkerProgressChanged2
        AddHandler BGWorkMode(g).RunWorkerCompleted, AddressOf WorkerCompleted2
        BGWorkMode(g).RunWorkerAsync()
        If Not BGWorkMode(g).IsBusy Then
            BGWorkMode(g).Dispose()
            g = 0
        End If
    End Sub

    Private Sub WorkerDoWork2(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        AuBkTimer01.Enabled = False
        ExportModeA()
    End Sub

    Private Sub WorkerProgressChanged2(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted2(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        'AuBkBtn05.Enabled = True
        AuBkTimer01.Enabled = True

    End Sub
#End Region

    Dim Release10SecMemo As Int32

    Private Sub AuBkTimer01_Tick(sender As Object, e As EventArgs) Handles AuBkTimer01.Tick
        AuBkDP03.Text = Now.ToString("hh:mm:ss tt")
        AuBkDP04.Text = Now.ToString("hh:mm:ss tt")

        If AuBkDP03.Text = AuBkDP01.Text Then
            'ExportModeA()
            OnClickTheWorker2()
        End If

        If AuBkDP04.Text = AuBkDP02.Text Then
            'ExportModeA()
            OnClickTheWorker()
        End If

        Release10SecMemo += 1

        Select Case Release10SecMemo
            Case Is >= 10800
                FlushMemory()
                Release10SecMemo = 0
                Me.Refresh()

        End Select

    End Sub

    Sub SendtoEmail()

        Try
            Dim SmtpServer As New SmtpClient("localhost", 587)
            Dim mail As New MailMessage()
            SmtpServer.Credentials = New Net.NetworkCredential("allan@uni-latex.com", "allan")
            SmtpServer.EnableSsl = True
            SSLValidation()
            mail = New MailMessage()
            mail.IsBodyHtml = True
            SmtpServer.Host = "smtp.uni-latex.com"
            mail.From = New MailAddress("allan@uni-latex.com")
            mail.To.Add("hodiumalchest@gmail.com")

            Dim attachment As System.Net.Mail.Attachment
            attachment = New System.Net.Mail.Attachment(AttachmentForMail)
            mail.Attachments.Add(attachment)

            mail.Subject = "ZIP File Back Up System"
            mail.Body = " " & vbCrLf & "<br />"
            mail.Body += " Dear Admin" & vbCrLf & "<br />"
            mail.Body += " This is your ZIP File Database BackUp System, " & vbCrLf & "<br />"
            mail.Body += " Please Keep it as your BackUp  " & vbCrLf & "<br />"
            mail.Body += " " & vbCrLf & "<br />"
            mail.Body += " " & vbCrLf & "<br />"
            mail.Body += " " & vbCrLf & "<br />"
            mail.Body += " From IT Department "
            mail.IsBodyHtml = True
            SmtpServer.Send(mail)
            mail.Dispose()

        Catch ex As Exception

        End Try
    End Sub
    Sub SSLValidation()
        System.Net.ServicePointManager.ServerCertificateValidationCallback = Function(s, cert, chain, sslPolicyErrors)
                                                                                 Return True
                                                                             End Function
    End Sub

    Private Sub AuBkBtn06_Click(sender As Object, e As EventArgs) Handles AuBkBtn06.Click
        Process.Start("explorer.exe", Application.StartupPath + "\Database List")
    End Sub

    Private Sub AuBkHideMe_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles AuBkHideMe.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        AuBkHideMe.Visible = True
    End Sub

    Private Sub MainAutoBk_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            AuBkHideMe.Visible = True
            Me.Hide()
        End If
    End Sub

    Private Sub AuBkChk01_CheckedChanged(sender As Object, e As EventArgs) Handles AuBkChk01.CheckedChanged
        Select Case AuBkChk01.Checked

            Case True
                AuBkChk01.Checked = True

            Case False
                AuBkChk01.Checked = True

        End Select
    End Sub

    Private Sub AuBkGrid01_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles AuBkGrid01.CellContentClick

    End Sub

    Private Sub AuBkGrid01_SelectionChanged(sender As Object, e As EventArgs) Handles AuBkGrid01.SelectionChanged

        Select Case GetKeyNow
            Case False
                AuBkGrid01.ClearSelection()

        End Select
    End Sub

    Private Sub AuBk07_KeyPress(sender As Object, e As KeyPressEventArgs) Handles AuBk07.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            Select Case AuBk07.Text
                Case AdminKeyAccept

                    Select Case GetKeyNow
                        Case False
                            GetKeyNow = True
                        Case True
                            GetKeyNow = False
                    End Select

                Case Else
                    GetKeyNow = False

            End Select
            AuBk07.Clear()
            AntiPeeping()
            e.Handled = True
        End If
    End Sub

    Private Sub SizeBtn02_Click(sender As Object, e As EventArgs) Handles SizeBtn02.Click

        Me.Size = New Drawing.Size(420, 150)
        TabControl1.TabPages(0).Enabled = False
        TabControl1.TabPages(1).Enabled = False
        TabControl1.TabPages(2).Enabled = False

    End Sub

    Private Sub SizeBtn01_Click(sender As Object, e As EventArgs) Handles SizeBtn01.Click
        Me.Size = New Drawing.Size(779, 594)

        TabControl1.TabPages(0).Enabled = True
        TabControl1.TabPages(1).Enabled = True
        TabControl1.TabPages(2).Enabled = True

    End Sub
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Me.Dispose()
    End Sub

    Private Sub MemoModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MemoModeToolStripMenuItem.Click
        Me.Size = New Drawing.Size(410, 150)
        TabControl1.TabPages(0).Enabled = False
        TabControl1.TabPages(1).Enabled = False
        TabControl1.TabPages(2).Enabled = False
    End Sub

    Private Sub NormalModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalModeToolStripMenuItem.Click
        Me.Size = New Drawing.Size(779, 594)

        TabControl1.TabPages(0).Enabled = True
        TabControl1.TabPages(1).Enabled = True
        TabControl1.TabPages(2).Enabled = True


    End Sub


  
End Class
