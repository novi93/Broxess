Imports System.Configuration
Imports System.IO
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Class frmSetupV5

    Inherits MetroWindow

#Region "model"
    Protected xSource As String
    Protected xDev As String
    Protected xName As String = "SETUP_V5_"
#End Region

#Region "method"
    Private Function createPG(path As String, pgCode As String, ByRef ErrMsg As String, Optional isAppend As Boolean = False) As Boolean
        Try
            ErrMsg = String.Empty
            Dim request = New RequestInfo With {
                .ProgramCode_Source = pgCode
                }
            Dim targetInfo = New TargetInfo With {
                .ProgramCode = pgCode
            }

            Dim outputInfo = New OutputInfo With {
                .Path = System.IO.Path.Combine(path, targetInfo.ProgramCode),
                .MvcFodler = "MVC",
                .ProxyFolder = "Proxy",
                .ProcessFolder = "Process"
                }

            If isAppend Then
                outputInfo.Path = System.IO.Path.Combine(path)
            End If
            Dim sourceData = DbHelper.LoadData(request)

            Dim sqlData = New SqlTextData With {
                .MvcData = SqlGenerator.GenerateScriptDs(sourceData.MvcData),
                .ProxyData = SqlGenerator.GenerateScriptDs(sourceData.ProxyData),
                .ProcessData = SqlGenerator.GenerateScriptDs(sourceData.ProcessData)
            }
            FileHelper.OutputToiawase(sqlData, outputInfo, isAppend)

            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function


    Private Sub StartProcess(Optional isAppend As Boolean = False)
        'checkError()
        'Dim connectInfo As New ConnectionInfo With {
        '    .mvc = Trim(Me.txtMvc.Text),
        '    .proxy = Trim(Me.txtProxy.Text),
        '    .process = Trim(Me.txtProcess.Text)
        '    }

        'DbHelper.Init(connectInfo)
    End Sub

    Private Sub FinishProcess()

    End Sub
#End Region

    Private Sub checkError()
        'Dim textRange As TextRange = New TextRange(txtPG.Document.ContentStart, txtPG.Document.ContentEnd)
        'Dim source = Trim(textRange.Text)
        'Dim path = Trim(Me.txtPath.Text)
        'Dim process = Trim(Me.txtProcess.Text)
        'Dim proxy = Trim(Me.txtProxy.Text)
        'Dim mvc = Trim(Me.txtMvc.Text)
        'If String.IsNullOrWhiteSpace(source) Then
        '    Throw New Exception("Source PG Code is Empty")
        'End If

        'If String.IsNullOrWhiteSpace(path) Then
        '    Throw New Exception("Path is Empty")
        'End If
        'Dim di = New DirectoryInfo(path)

        'If String.IsNullOrWhiteSpace(process) Then
        '    Throw New Exception("process connection is Empty")
        'End If
        'If String.IsNullOrWhiteSpace(proxy) Then
        '    Throw New Exception("proxy connection is Empty")
        'End If
        'If String.IsNullOrWhiteSpace(mvc) Then
        '    Throw New Exception("mvc connection is Empty")
        'End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Owner.Activate()
    End Sub

    Private Sub frmSetupV5_Closed(sender As Object, e As EventArgs) Handles Me.Closing
       
    End Sub

    Protected Sub saveSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        config.AppSettings.Settings(xName & "_txtSource").Value = Trim(Me.txtSource.Text)
        config.AppSettings.Settings(xName & "_txtDevenv2k5").Value = Trim(Me.txtDevenv2k5.Text)
        config.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("appSettings")
    End Sub
    Private Sub frmSetupV5_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        loadSetting()
    End Sub
    Protected Sub loadSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Me.txtSource.Text = config.AppSettings.Settings(xName & "txtSource").Value
        Me.txtDevenv2k5.Text = config.AppSettings.Settings(xName & "txtDevenv2k5").Value
        Me.Title = String.Format("{0} [{1}]",
                         Me.Title,
                        config.AppSettings.Settings("Version").Value)
    End Sub
    'Private Async Function Export(Optional isAppend As Boolean = False) As Task(Of Integer)

    '    'Dim isSuccess As Boolean = False
    '    'Dim xMsgResult = String.Empty
    '    'Dim controller = Await ShowProgressAsync("Please wait", "Check Error")
    '    'Try

    '    '    StartProcess()
    '    '    Dim textRange As TextRange = New TextRange(txtPG.Document.ContentStart, txtPG.Document.ContentEnd)
    '    '    Dim listPg = textRange.Text.Trim
    '    '    Dim seprator As String() = Nothing
    '    '    Dim lst = listPg.Split(seprator, StringSplitOptions.RemoveEmptyEntries)
    '    '    Array.Sort(lst)

    '    '    Dim xMsg As String = String.Empty
    '    '    Dim cnt As Integer = lst.Count
    '    '    Dim xPath = txtPath.Text.Trim
    '    '    Dim isExist As Boolean = Directory.Exists(xPath)

    '    '    Dim isNotEmpty As Boolean = isExist AndAlso Not Directory.EnumerateFiles(xPath).Any()

    '    '    If isAppend AndAlso isExist AndAlso isNotEmpty Then
    '    '        If Await Me.ShowMessageAsync("Delete Foder !",
    '    '                                     "Folder 「" & xPath & "」 Will Be DELETED Before Export?",
    '    '            MessageDialogStyle.AffirmativeAndNegative, New MetroDialogSettings With {.DefaultButtonFocus = MessageDialogResult.Negative}
    '    '                                     ) = MessageDialogResult.Affirmative Then
    '    '            Dim di As New DirectoryInfo(xPath)

    '    '            For Each fi In di.GetFiles()
    '    '                fi.Delete()
    '    '            Next

    '    '            For Each sdi In di.GetDirectories()
    '    '                sdi.Delete(True)
    '    '            Next
    '    '        Else
    '    '            Await controller.CloseAsync()
    '    '            Return 0
    '    '        End If
    '    '    End If

    '    '    xMsg = String.Format("Processing... {0} program(s)", cnt)
    '    '    controller.SetMessage(xMsg)
    '    '    controller.SetIndeterminate()
    '    '    If Not Directory.Exists(xPath) Then
    '    '        Directory.CreateDirectory(xPath)
    '    '    End If

    '    '    Dim tskRun As Task(Of Integer) = Task.Run(
    '    '        Function()
    '    '            Dim cntCurrent = 0

    '    '            For i = 0 To lst.Count - 1
    '    '                Dim xPgCode = lst(i)

    '    '                cntCurrent = cntCurrent + 1
    '    '                xMsg = String.Format("Processing... {0,3} / {1,-3} Program(s) [{2}]", cntCurrent, cnt, xPgCode)
    '    '                controller.SetMessage(xMsg)


    '    '                If createPG(xPath, xPgCode, xMsg, isAppend) Then
    '    '                    xMsgResult &= vbNewLine & String.Format("※ {0,-12} ", xPgCode)
    '    '                Else
    '    '                    xMsgResult &= vbNewLine & String.Format("※ {0,-12} ERROR: {1} ", xPgCode, xMsg)
    '    '                End If
    '    '            Next

    '    '            Return 1
    '    '        End Function)

    '    '    '// run tsk
    '    '    Dim tsk As Integer = Await tskRun
    '    '    isSuccess = (tsk = 1)
    '    'Catch ex As Exception
    '    '    isSuccess = False
    '    '    xMsgResult = ex.Message()
    '    'End Try
    '    ' ''Close...
    '    'Await controller.CloseAsync()
    '    'If isSuccess Then
    '    '    If Await Me.ShowMessageAsync("Done!", xMsgResult & vbNewLine & vbNewLine &
    '    '                                 "Open output Folder?",
    '    '                                 MessageDialogStyle.AffirmativeAndNegative, New MetroDialogSettings With {.DefaultButtonFocus = MessageDialogResult.Affirmative}
    '    '                                 ) = MessageDialogResult.Affirmative Then
    '    '        System.Diagnostics.Process.Start(Trim(txtPath.Text))
    '    '    End If
    '    'Else
    '    '    Await Me.ShowMessageAsync(GlobalSetting.ERROR_MESSAGE_TITLE, xMsgResult, MessageDialogStyle.Affirmative)
    '    'End If
    '    Return 0
    'End Function

    Private Sub btnbuild_Click(sender As Object, e As RoutedEventArgs) Handles btnbuild.Click
        xSource = txtSource.Text
        xDev = txtDevenv2k5.Text

        Dim xParam As String

        xParam = String.Format("""{0}"" /Build \""Release|x86""", xSource)
        Utils.RunCommandCom(xDev, xParam, False)

    End Sub
End Class
