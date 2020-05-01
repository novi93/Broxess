Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.Configuration

Class frmDuplicate
    Inherits MetroWindow

    Private Async Sub Button_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        Await ExportAsync()

    End Sub

#Region "method"
    Private Sub StartProcess()
        checkError()
        Dim connectInfo As New ConnectionInfo With {
           .mvc = Trim(Me.txtMvc.Text),
           .proxy = Trim(Me.txtProxy.Text),
           .process = Trim(Me.txtProcess.Text)
           }

        DbHelper.Init(connectInfo)
    End Sub

    Private Sub FinishProcess(info As OutputInfo)

    End Sub
#End Region

    Private Sub checkError()

        Dim source = Trim(Me.txtSourceCode.Text)
        Dim target = Trim(Me.txtTgCode.Text)
        Dim path = Trim(Me.txtPath.Text)
        If String.IsNullOrWhiteSpace(source) Then
            Throw New Exception("Source PG Code is Empty")
        End If
        If String.IsNullOrWhiteSpace(target) Then
            Throw New Exception("Target PG Code is Empty")
        End If
        If Not Date.TryParse(Me.txtDate.Text, Nothing) Then
            Throw New Exception("Cannot parse AD_DATE")
        End If
        If String.IsNullOrWhiteSpace(path) Then
            Throw New Exception("Path is Empty")
        End If
        '' check valid path
        Dim di = New DirectoryInfo(path)

        If Not Directory.Exists(path) Then
            Directory.CreateDirectory(path)
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Owner.Activate()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        LoadSetting()
        Me.txtSourceCode.Focus()

        Me.chkGenSource.Visibility = System.Windows.Visibility.Hidden
        Me.chkAddPending.Visibility = System.Windows.Visibility.Hidden

    End Sub

    Private Sub MetroWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        SaveSetting()
    End Sub

    Private Async Function ExportAsync() As Task(Of Integer)

        Dim isSuccess As Boolean = False
        Dim xMsgResult = String.Empty
        Dim controller = Await ShowProgressAsync("Please wait", String.Empty)
        controller.SetIndeterminate()
        Try
            controller.SetMessage("Check Error")
            StartProcess()
            Dim xMsg As String = String.Empty
            Dim xPath = txtPath.Text.Trim

            xMsg = String.Format("Exporting")
            controller.SetMessage(xMsg)

            Dim request = New RequestInfo With {
                        .ProgramCode_Source = Me.txtSourceCode.Text,
                        .ProgramCode_Target = Me.txtTgCode.Text,
                        .UserNo = 0
                        }
            Dim targetInfo = New TargetInfo With {
                .ProgramCode = Me.txtTgCode.Text,
                .ProgramNameKanji = Me.txtTgName1.Text,
                .ProgramNameRomanji = Me.txtTgName2.Text,
                .ProgramOldCode = Me.txtSourceCode.Text,
                .ProgramOldName = "OldName",
                .UserNo = 0,
                .AdDate = Date.Parse(Me.txtDate.Text),
                .AdUser = Me.txtUser.Text,
                .AdOPID = Me.txtOpid.Text,
                .UdDate = Date.Parse(Me.txtDate.Text),
                .UdUser = Me.txtUser.Text,
                .UdOPID = Me.txtOpid.Text,
                .Guid = Guid.NewGuid
                }
            Dim outputInfo = New OutputInfo With {
                .Path = System.IO.Path.Combine(Me.txtPath.Text),
                .MvcFodler = "MVC",
                .ProxyFolder = "Proxy",
                .ProcessFolder = "Process",
                .IsOutputKengen = Me.chkKengen.IsChecked,
                .IsOutputMenu = Me.chkMenu.IsChecked,
                .IsUpdateDB = Me.chkUpdateDB.IsChecked,
                .IsGenSource = Me.chkGenSource.IsChecked,
                .IsAddPendding = Me.chkAddPending.IsChecked
                }
            Dim tskRun As Task(Of Integer) = Task.Run(
                Function()

                    Dim sourceData = DbHelper.LoadData(request)
                    Dim outputData = DataHelper.CustomizeData(sourceData, targetInfo)

                    ''todo 
                    If Not outputInfo.IsOutputMenu Then
                        outputData.MvcData.Tables.Remove(TableName.Menu)
                        outputData.ProxyData.Tables.Remove(TableName.T_MENU_EXTENSION)
                    End If

                    If Not outputInfo.IsOutputKengen Then
                        outputData.ProcessData.Tables.Remove(TableName.M_PROGRAM_ALLOW)
                    End If

                    Dim sqlData = New SqlTextData With {
                        .MvcData = SqlGenerator.GenerateScriptDs(outputData.MvcData),
                        .ProxyData = SqlGenerator.GenerateScriptDs(outputData.ProxyData),
                        .ProcessData = SqlGenerator.GenerateScriptDs(outputData.ProcessData)
                    }
                    FileHelper.Output(sqlData, outputInfo)
                    If outputInfo.IsUpdateDB Then
                        DbHelper.UpdateDB(outputInfo)
                    End If
                    If outputInfo.IsGenSource Then
                    End If

                    If outputInfo.IsAddPendding Then
                    End If

                    Me.FinishProcess(outputInfo)

                    Return 1
                End Function)

            '// run tsk
            Dim tsk As Integer = Await tskRun
            isSuccess = (tsk = 1)
        Catch ex As Exception
            isSuccess = False
            xMsgResult = ex.Message()
        End Try
        ''Close...
        Await controller.CloseAsync()
        If isSuccess Then
            Dim msg As String = ""
            If Me.chkUpdateDB.IsChecked Then
                msg = xMsgResult & vbNewLine & vbNewLine & "Update DB Done!! " & vbNewLine & "Open output Folder?"
            Else
                msg = xMsgResult & vbNewLine & vbNewLine & "Open output Folder?"
            End If

            If Await Me.ShowMessageAsync("Done!", msg,
                                         MessageDialogStyle.AffirmativeAndNegative, New MetroDialogSettings With {.DefaultButtonFocus = MessageDialogResult.Affirmative}
                                         ) = MessageDialogResult.Affirmative Then
                System.Diagnostics.Process.Start(Trim(txtPath.Text))
            End If
        Else
            Await Me.ShowMessageAsync(GlobalSetting.ERROR_MESSAGE_TITLE, xMsgResult, MessageDialogStyle.Affirmative)
        End If
        Return 0
    End Function

    Public Sub SaveSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

        config.AppSettings.Settings("DUP_Path").Value = Trim(Me.txtPath.Text)
        config.AppSettings.Settings("DUP_txtSourceCode").Value = Trim(Me.txtSourceCode.Text)
        config.AppSettings.Settings("DUP_txtTgCode").Value = Trim(Me.txtTgCode.Text)
        config.AppSettings.Settings("DUP_txtTgName1").Value = Trim(Me.txtTgName1.Text)
        config.AppSettings.Settings("DUP_txtTgName2").Value = Trim(Me.txtTgName2.Text)
        config.AppSettings.Settings("DUP_txtDate").Value = Trim(Me.txtDate.Text)
        config.AppSettings.Settings("DUP_txtUser").Value = Trim(Me.txtUser.Text)
        config.AppSettings.Settings("DUP_txtOpid").Value = Trim(Me.txtOpid.Text)
        config.AppSettings.Settings("DUP_chkMenu").Value = Me.chkMenu.IsChecked
        config.AppSettings.Settings("DUP_chkKengen").Value = Me.chkKengen.IsChecked
        config.AppSettings.Settings("DUP_chkUpdateDB").Value = Me.chkUpdateDB.IsChecked
        config.AppSettings.Settings("DUP_chkGenSource").Value = Me.chkGenSource.IsChecked
        config.AppSettings.Settings("DUP_chkAddPending").Value = Me.chkAddPending.IsChecked

        config.ConnectionStrings.ConnectionStrings("DUP_ProcessConnection").ConnectionString = Trim(Me.txtProcess.Text)
        config.ConnectionStrings.ConnectionStrings("DUP_ProxyConnection").ConnectionString = Trim(Me.txtProxy.Text)
        config.ConnectionStrings.ConnectionStrings("DUP_UsoliaMvc").ConnectionString = Trim(Me.txtMvc.Text)
        config.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("appSettings")
    End Sub

    Public Sub LoadSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Me.txtPath.Text = config.AppSettings.Settings("DUP_Path").Value
        Me.txtSourceCode.Text = config.AppSettings.Settings("DUP_txtSourceCode").Value
        Me.txtTgCode.Text = config.AppSettings.Settings("DUP_txtTgCode").Value
        Me.txtTgName1.Text = config.AppSettings.Settings("DUP_txtTgName1").Value
        Me.txtTgName2.Text = config.AppSettings.Settings("DUP_txtTgName2").Value
        Me.txtDate.Text = config.AppSettings.Settings("DUP_txtDate").Value
        Me.txtUser.Text = config.AppSettings.Settings("DUP_txtUser").Value
        Me.txtOpid.Text = config.AppSettings.Settings("DUP_txtOpid").Value
        Me.chkMenu.IsChecked = config.AppSettings.Settings("DUP_chkMenu").Value
        Me.chkKengen.IsChecked = config.AppSettings.Settings("DUP_chkKengen").Value
        Me.chkUpdateDB.IsChecked = config.AppSettings.Settings("DUP_chkUpdateDB").Value
        Me.chkGenSource.IsChecked = config.AppSettings.Settings("DUP_chkGenSource").Value
        Me.chkAddPending.IsChecked = config.AppSettings.Settings("DUP_chkAddPending").Value
        Me.txtProcess.Text = config.ConnectionStrings.ConnectionStrings("DUP_ProcessConnection").ConnectionString
        Me.txtProxy.Text = config.ConnectionStrings.ConnectionStrings("DUP_ProxyConnection").ConnectionString
        Me.txtMvc.Text = config.ConnectionStrings.ConnectionStrings("DUP_UsoliaMvc").ConnectionString
        Me.Title = String.Format("{0} [{1}]",
                         Me.Title,
                        config.AppSettings.Settings("Version").Value)
    End Sub
End Class
