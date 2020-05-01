Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.Configuration
Imports System.Xml
Imports System.Collections
Imports System.Text
Imports System.Linq

Class frmDuplicateSearch
    Inherits MetroWindow

    'Private Async Sub Button_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
    '    Await ExportAsync()

    'End Sub

    Enum eAddMode
        EOF
        AboveLargerPG
    End Enum

    Protected filePath As String
    Protected doc As XmlDocument
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        doc = New XmlDocument With {
                                        .PreserveWhitespace = True
                                    }
    End Sub
#Region "method"
    Private Sub StartProcess()
    End Sub

    Private Sub FinishProcess(info As OutputInfo)

    End Sub
#End Region

    Private Sub checkError()
        filePath = txtPath.Text
        Dim path = Trim(Me.txtPath.Text)
        If String.IsNullOrWhiteSpace(path) Then
            Throw New Exception("Path is Empty")
        End If
        ' '' check valid path
        Dim di = New FileInfo(path)
        If di.Extension.Equals("xml") Then
            Throw New Exception("Please input *.xml file")
        End If

    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        LoadSetting()

    End Sub

    Private Sub MetroWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        SaveSetting()
    End Sub

    Public Sub SaveSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

        config.AppSettings.Settings("Search_Path").Value = Trim(Me.txtPath.Text)
        config.AppSettings.Settings("Search_txtOldCode").Value = Trim(Me.txtOldCode.Text)
        config.AppSettings.Settings("Search_txtNewCode").Value = Trim(Me.txtNewCode.Text)
        config.AppSettings.Settings("Search_radContinuePG").Value = Trim(Me.radContinuePG.IsChecked)
        config.AppSettings.Settings("Search_radEndfile").Value = Trim(Me.radEndfile.IsChecked)
        config.AppSettings.Settings("Search_txtPGFrom").Value = Trim(Me.txtPGFrom.Text)
        config.AppSettings.Settings("Search_txtPGTo").Value = Trim(Me.txtPGTo.Text)

        config.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("appSettings")
    End Sub

    Public Sub LoadSetting()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Me.txtPath.Text = config.AppSettings.Settings("Search_Path").Value
        Me.txtOldCode.Text = config.AppSettings.Settings("Search_txtOldCode").Value
        Me.txtNewCode.Text = config.AppSettings.Settings("Search_txtNewCode").Value
        Me.radContinuePG.IsChecked = config.AppSettings.Settings("Search_radContinuePG").Value
        Me.radEndfile.IsChecked = config.AppSettings.Settings("Search_radEndfile").Value
        Me.txtPGFrom.Text = config.AppSettings.Settings("Search_txtPGFrom").Value
        Me.txtPGTo.Text = config.AppSettings.Settings("Search_txtPGTo").Value

        Me.Title = String.Format("{0} [{1}]",
                         Me.Title,
                        config.AppSettings.Settings("Version").Value)
    End Sub

    Private Sub btnLoadXml_Click(sender As Object, e As RoutedEventArgs) Handles btnLoadXml.Click
        checkError()
        doc.Load(filePath)
        Dim root As XmlNode = doc.DocumentElement
        Dim nodeList As XmlNodeList
        Dim newCode As String = Me.txtNewCode.Text
        Dim oldCode As String = Me.txtOldCode.Text.Trim
        Dim query As String
        Dim sb As New StringBuilder(root.OuterXml)
        sb.Clear()
        If String.IsNullOrWhiteSpace(oldCode) Then
            query = String.Empty
            sb.AppendLine(doc.OuterXml)
        Else
            query = String.Format("descendant::Table[SQL_CODE={0}]", oldCode)
            nodeList = root.SelectNodes(query)

            Dim aaa = nodeList.Cast(Of XmlNode)()
            For Each node As XmlNode In aaa
                node.SelectSingleNode("SQL_CODE").InnerText = newCode
                sb.AppendLine(node.OuterXml)
            Next
        End If
        txtPreview.Clear()
        txtPreview.Text = sb.ToString

    End Sub

    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs) Handles btnClear.Click
        txtPreview.Clear()
    End Sub

    Private Sub btnAddToXML_Click(sender As Object, e As RoutedEventArgs) Handles btnAddToXML.Click
        Dim mode As eAddMode
        If radEndfile.IsChecked Then
            mode = eAddMode.EOF
        End If

        checkError()
        doc.Load(filePath)
        Dim oldCode As String = Me.txtOldCode.Text.Trim
        Dim query As String
        query = String.Format("descendant::Table[SQL_CODE={0}]", oldCode)
        Dim newdoc As New XmlDocument With {.PreserveWhitespace = True}
        newdoc.LoadXml(String.Format("<dummyRoot>{0}{1}{2}</dummyRoot>", Environment.NewLine, txtPreview.Text, Environment.NewLine))
        Dim newnodes = newdoc.ChildNodes.Item(0).ChildNodes

        'Dim xd As XDocument = XDocument.Parse(txtPreview.Text)
        'Dim newNodes = xd.Document.Elements


        Dim node = doc.SelectSingleNode(query)
        Dim aaa = newnodes.Cast(Of XmlNode)()

        If mode = eAddMode.EOF OrElse node Is Nothing Then
            ''add  to end
            For i As Integer = newnodes.Count - 1 To 0 Step -1
                doc.ImportNode(newnodes.ItemOf(i), True)
                'doc.AppendChild(newnodes.ItemOf(i))
            Next

        Else
            '' add above
        End If

    End Sub
End Class
