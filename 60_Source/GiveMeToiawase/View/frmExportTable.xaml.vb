Imports System.Configuration
Imports System.Data
Imports System.IO
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs

Public Class frmExportTable
    Inherits MetroWindow

    Dim _tableInfoHelper As New TableInfoHelper()

    'Private config As Configuration


    Public Sub New()
        DataContext = Me

        InitializeComponent()

    End Sub

#Region "Events"
    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Owner.Activate()
    End Sub

    Private Sub frmExport_Loaded(sender As Object, e As RoutedEventArgs)

        LoadSetting()

        grdMain.ItemsSource = _tableInfoHelper.getTableInfoList()

    End Sub

    Private Sub frmExport_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles frmExport.Closing
        SaveSetting()
    End Sub

    Private Async Sub btnExport_Click(sender As Object, e As RoutedEventArgs) Handles btnExport.Click
        Await Export()
    End Sub
#End Region

#Region "Method"

    Private Sub CheckError()
        Dim xServer = Trim(Me.txtServer.Text)
        Dim path = Trim(Me.txtPath.Text)
        If String.IsNullOrWhiteSpace(xServer) Then
            Throw New Exception("Server Empty")
        End If

        If String.IsNullOrWhiteSpace(path) Then
            Throw New Exception("Path is Empty")
        End If
        Dim di = New DirectoryInfo(path)
    End Sub

    Private Sub SaveSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        config.AppSettings.Settings("frmExport_Server").Value = Me.txtServer.Text.Trim
        config.AppSettings.Settings("frmExport_outputPath").Value = Me.txtPath.Text.Trim
        config.Save(ConfigurationSaveMode.Minimal)
        ConfigurationManager.RefreshSection("appSettings")

        _tableInfoHelper.SaveSetting()
    End Sub

    Private Sub LoadSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Me.txtServer.Text = config.AppSettings.Settings("frmExport_Server").Value
        Me.txtPath.Text = config.AppSettings.Settings("frmExport_outputPath").Value
        Me.Title = String.Format("{0} [{1}]",
                         Me.Title,
                        config.AppSettings.Settings("Version").Value)
        _tableInfoHelper.LoadSetting()
    End Sub

    Private Async Function Export() As Task(Of Integer)

        If Not _tableInfoHelper.isEmptyOrWhiteSpace() Then

            Dim isSuccess As Boolean = False
            Dim xMsgResult = String.Empty
            Dim controller = Await ShowProgressAsync("Please wait", "Check Error")
            Try
                CheckError()
                Dim xMsg As String = String.Empty
                Dim cnt As Integer = _tableInfoHelper.Count
                Dim _tableInfo = _tableInfoHelper.getTableInfoList()
                Dim repo As New Repository(txtServer.Text.Trim)
                Dim xPath = txtPath.Text.Trim

                xMsg = String.Format("Exporting {0} table(s)...", cnt)
                controller.SetMessage(xMsg)
                controller.SetIndeterminate()
                If Not Directory.Exists(xPath) Then
                    Directory.CreateDirectory(xPath)
                End If

                Dim tskRun As Task(Of Integer) = Task.Run(
                    Function()
                        Dim cntCurrent = 0
                        Dim tInfo As TableInfoEntity
                        Dim xTableName As String
                        Dim xCondition As String
                        Dim xOrder As String
                        Dim dt As DataTable

                        For i = 0 To _tableInfo.Count - 1
                            tInfo = _tableInfo.Item(i)
                            xTableName = tInfo.TableName
                            xCondition = tInfo.Condition
                            xOrder = tInfo.Order
                            If String.IsNullOrWhiteSpace(xTableName) Then
                                Continue For
                            End If

                            xTableName = DbHelper.GetCorrectTableName(xTableName, repo)
                            tInfo.TableName = xTableName

                            cntCurrent = cntCurrent + 1
                            xMsg = String.Format("Exporting... {0,3} / {1,-3} table(s) [{2}]", cntCurrent, cnt, xTableName)
                            controller.SetMessage(xMsg)

                            dt = repo.GetData(tInfo.createQuerry()).Copy
                            dt.TableName = xTableName

                            If FileHelper.Output(SqlGenerator.GenerateScript(dt), xPath) Then
                                xMsgResult &= vbNewLine & String.Format("※ {0,-40}:{1,10} row(s)", xTableName, dt.Rows.Count)
                            End If
                        Next

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
                If Await Me.ShowMessageAsync("Done!", xMsgResult & vbNewLine & vbNewLine &
                                             "Open output Folder?",
                                             MessageDialogStyle.AffirmativeAndNegative, New MetroDialogSettings With {.DefaultButtonFocus = MessageDialogResult.Affirmative}
                                             ) = MessageDialogResult.Affirmative Then
                    System.Diagnostics.Process.Start(Trim(txtPath.Text))
                End If
            Else
                Await Me.ShowMessageAsync(GlobalSetting.ERROR_MESSAGE_TITLE, xMsgResult, MessageDialogStyle.Affirmative)
            End If
        End If

        Return 0
    End Function

#End Region

End Class
