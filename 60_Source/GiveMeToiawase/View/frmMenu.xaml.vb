Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Threading
Imports System.Windows.Controls.Primitives
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports MaterialDesignThemes.Wpf

Public Class frmMenu
    Inherits Window

    Public Shared Snackbar As Snackbar = New Snackbar()

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Task.Factory.StartNew(Sub() Thread.Sleep(2500)).ContinueWith(Sub(t)
                                                                         '// note you can use the message queue from any thread, but just for the demo here we 
                                                                         '//need to get the message queue from the snackbar, so need to be on the dispatcher
                                                                         MainSnackbar.MessageQueue?.Enqueue("Welcome to Brocess")
                                                                     End Sub, TaskScheduler.FromCurrentSynchronizationContext())
        ''todo
        'DataContext = New MainWindowViewModel(MainSnackbar.MessageQueue!)

        Dim paletteHelper = New PaletteHelper()
        Dim Theme = paletteHelper.GetTheme()

        DarkModeToggleButton.IsChecked = (Theme.GetBaseTheme() = BaseTheme.Dark)
        Dim themeManager = paletteHelper.GetThemeManager()
        If (themeManager Is Nothing) Then
            AddHandler themeManager.ThemeChanged, Sub(o, e) DarkModeToggleButton.IsChecked = e.NewTheme?.GetBaseTheme() = BaseTheme.Dark
        End If

        Snackbar = MainSnackbar

    End Sub



    Private Sub UIElement_OnPreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
        ''until we had a StaysOpen glag to Drawer, this will help with scroll bars
        Dim dependencyObject = DirectCast(Mouse.Captured, DependencyObject)

        While (dependencyObject Is Nothing)

            If dependencyObject Is GetType(ScrollBar) Then
                Return
            End If

            dependencyObject = VisualTreeHelper.GetParent(dependencyObject)
        End While

        MenuToggleButton.IsChecked = False
    End Sub



    Private Async Sub MenuPopupButton_OnClick(sender As Object, e As RoutedEventArgs)
        '//todo
        'var sampleMessageDialog = New SampleMessageDialog
        '    {
        '        Message = {Text = ((ButtonBase)sender).Content.ToString() }
        '    };

        '    Await DialogHost.Show(sampleMessageDialog, "RootDialog");
    End Sub


    Private Sub OnCopy(sender As Object, e As RoutedEventArgs)

    End Sub
    Private Sub MenuToggleButton_OnClick(sender As Object, e As RoutedEventArgs)
        DemoItemsSearchBox.Focus()
    End Sub

    Private Sub MenuDarkModeButton_Click(sender As Object, e As RoutedEventArgs)
        ModifyTheme(DarkModeToggleButton.IsChecked)
    End Sub

    Private Shared Sub ModifyTheme(isDarkTheme As Boolean)
        Dim paletteHelper = New PaletteHelper()
        Dim Mytheme = paletteHelper.GetTheme()

        Mytheme.SetBaseTheme(If(isDarkTheme, MaterialDesignThemes.Wpf.Theme.Dark, MaterialDesignThemes.Wpf.Theme.Light))
        paletteHelper.SetTheme(Mytheme)
    End Sub

    'Private Sub Tile_Click(sender As Object, e As RoutedEventArgs)
    '    Dim frm As New frmExportTable()
    '    frm.Owner = Me
    '    frm.Show()
    'End Sub

    'Private Sub tltCopyToiMSS_Click(sender As Object, e As RoutedEventArgs) Handles tltCopyToiMSS.Click
    '    Dim frm As New frmExportToiawase()
    '    frm.Owner = Me
    '    frm.Show()
    'End Sub

    'Private Sub tltCopyToiKMT_Click(sender As Object, e As RoutedEventArgs) Handles tltCopyToiKMT.Click
    '    Dim frm As New frmDuplicate()
    '    frm.Owner = Me
    '    frm.Show()
    'End Sub

    'Private Sub tltDuptoiwase_Click(sender As Object, e As RoutedEventArgs) Handles tltDuptoiwase.Click
    '    Dim xMsg = "Đã nói là comingsoon rồi còn chày cối mà bấm vào >.< "
    '    'Me.ShowMessageAsync("Lỳ....", xMsg)
    '    MessageBox.Show("Lỳ....", xMsg)
    'End Sub

    'Private Sub frm_Loaded(sender As Object, e As RoutedEventArgs) Handles frmExport.Loaded
    '    Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
    '    Me.Title = String.Format("{0} [{1}]",
    '                             Me.Title,
    '                            config.AppSettings.Settings("Version").Value)

    '    Me.tltSettingSearch_Dup.Visibility = Visibility.Hidden
    '    Me.tltInitData.Visibility = Visibility.Hidden
    '    Me.tltSetupV5.Visibility = Visibility.Hidden

    'End Sub

    'Private Sub tltSettingSearch_Dup_Click(sender As Object, e As RoutedEventArgs) Handles tltSettingSearch_Dup.Click
    '    Dim frm As New frmDuplicateSearch()
    '    frm.Owner = Me
    '    frm.Show()
    'End Sub

    'Private Sub tltInitData_Click(sender As Object, e As RoutedEventArgs) Handles tltInitData.Click
    '    Dim frm As New frmInitData()
    '    frm.Owner = Me
    '    frm.Show()
    'End Sub

    'Private Sub tltSetupV5_Click(sender As Object, e As RoutedEventArgs) Handles tltSetupV5.Click
    '    Dim frm As New frmSetupV5()
    '    frm.Owner = Me
    '    frm.Show()
    'End Sub
End Class
