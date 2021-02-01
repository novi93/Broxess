Imports System.Reflection
Imports System
Imports System.Deployment.Application
Imports System.Configuration

Public Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Public Sub New()
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

        config.AppSettings.Settings("Version").Value = VersionLabel()
        config.Save()
    End Sub

    Public Function VersionLabel() As String
        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim ver = ApplicationDeployment.CurrentDeployment.CurrentVersion
            'Return String.Format("{4} [{0}.{1}.{2}.{3}]", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name)
            Return $"{ver.Major}.{ver.Minor}.{ ver.Build}.{ver.Revision}"
        Else
            Dim ver = Assembly.GetExecutingAssembly().GetName().Version
            'Return String.Format("{4} [{0}.{1}.{2}.{3}]", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name)
            Return $"{ver.Major}.{ver.Minor}.{ ver.Build}.{ver.Revision}"
        End If

    End Function
End Class
