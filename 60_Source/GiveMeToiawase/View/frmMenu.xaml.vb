Imports System.Configuration
Imports System.Data
Imports System.IO
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs

Public Class frmMenu
    Inherits MetroWindow

    Private Sub Tile_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New frmExportTable()
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub tltCopyToiMSS_Click(sender As Object, e As RoutedEventArgs) Handles tltCopyToiMSS.Click
        Dim frm As New frmExportToiawase()
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub tltCopyToiKMT_Click(sender As Object, e As RoutedEventArgs) Handles tltCopyToiKMT.Click
        Dim frm As New frmDuplicate()
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub tltDuptoiwase_Click(sender As Object, e As RoutedEventArgs) Handles tltDuptoiwase.Click
        Dim xMsg = "Đã nói là comingsoon rồi còn chày cối mà bấm vào >.< "
        Me.ShowMessageAsync("Lỳ....", xMsg)
    End Sub

    Private Sub frm_Loaded(sender As Object, e As RoutedEventArgs) Handles frmExport.Loaded
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Me.Title = String.Format("{0} [{1}]",
                                 Me.Title,
                                config.AppSettings.Settings("Version").Value)

        Me.tltSettingSearch_Dup.Visibility = Visibility.Hidden
        Me.tltInitData.Visibility = Visibility.Hidden
        Me.tltSetupV5.Visibility = Visibility.Hidden

    End Sub

    Private Sub tltSettingSearch_Dup_Click(sender As Object, e As RoutedEventArgs) Handles tltSettingSearch_Dup.Click
        Dim frm As New frmDuplicateSearch()
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub tltInitData_Click(sender As Object, e As RoutedEventArgs) Handles tltInitData.Click
        Dim frm As New frmInitData()
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub tltSetupV5_Click(sender As Object, e As RoutedEventArgs) Handles tltSetupV5.Click
        Dim frm As New frmSetupV5()
        frm.Owner = Me
        frm.Show()
    End Sub
End Class
