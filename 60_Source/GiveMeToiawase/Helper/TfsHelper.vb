'Imports Microsoft.TeamFoundation.Client
'Imports Microsoft.TeamFoundation.VersionControl.Client
'Imports System
'Imports System.Collections.Generic
'Imports System.IO
'Imports System.Linq
'Imports System.Reflection
'Imports System.Text
'Imports System.Threading.Tasks
'Imports System.Windows

'Public Class TfsHelper

'    Protected Shared baseIsConnected As Boolean = False
'    Protected Shared baseLogPath As String = Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName, "TFS_LOG")
'    Protected Shared baseLogWriter As StreamWriter = Nothing
'    Protected Shared baseTfsProjectCollection As TfsTeamProjectCollection
'    Protected Shared baseVersionControlServer As VersionControlServer
'    Protected Shared baseWorkspace As Workspace
'    Protected Shared baseTfsUrl As String = "https://tfs-dev3.its-process.net:8081/tfs/DefaultCollection"
'    Public Shared baseWorkspacePath As String = String.Empty
'    Public Shared baseNewPendingChangeLog As String = String.Empty

'    Public Shared Function ConnectToTfs() As Boolean
'        Try
'            If (baseIsConnected) Then
'                Return True
'            End If

'            If (Not Directory.Exists(baseLogPath)) Then
'                Directory.CreateDirectory(baseLogPath)
'            End If

'            baseTfsProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(New Uri(baseTfsUrl))
'            baseTfsProjectCollection.EnsureAuthenticated()

'            baseVersionControlServer = baseTfsProjectCollection.GetService(Of VersionControlServer)()

'            baseWorkspace = baseVersionControlServer.GetWorkspace(baseWorkspacePath)
'            If (Not baseWorkspace Is Nothing) Then
'                If (Not baseWorkspace.HasWorkspacePermission(WorkspacePermissions.CheckIn)) Then
'                    Throw New Exception("Error Check_WorkspacePermission [ CheckIn ] not availabel .!")
'                    Return False
'                End If
'            End If

'            baseIsConnected = True
'            Return True
'        Catch ex As Exception
'            'ModernDialog.ShowMessage(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButton.OK)
'            Throw
'            Return False
'        End Try

'    End Function


'    Public Shared Function PendingChangeToProject(sPath As String, Optional isRecursive As Boolean = True) As Boolean
'        If (Not PendingAdd(sPath, isRecursive)) Then
'            Return PendingEdit(sPath, isRecursive)
'        End If
'        Return True
'    End Function


'    Public Shared Function PendingAdd(sPath As String, Optional isRecursive As Boolean = True) As Boolean
'        Try
'            If (Not baseWorkspace Is Nothing) Then
'                Dim penCount As Integer = baseWorkspace.PendAdd(sPath, isRecursive)
'                Dim msg As String = "PendingAdd ( " + penCount + " ) :" + sPath
'                If (penCount > 0) Then
'                    AddLog(msg)
'                End If
'                Return penCount > 0
'            End If
'            Return False
'        Catch ex As Exception
'            'ModernDialog.ShowMessage(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButton.OK)
'            Throw
'            Return False
'        End Try

'    End Function

'    Public Shared Function PendingEdit(sPath As String, Optional isRecursive As Boolean = True) As Boolean
'        Try
'            If (Not baseWorkspace Is Nothing) Then
'                Dim penCount As Integer = baseWorkspace.PendEdit(sPath, If(isRecursive, RecursionType.Full, RecursionType.None))
'                Dim msg As String = "PendingEdit ( " + penCount + " ) :" + sPath
'                If (penCount > 0) Then
'                    AddLog(msg)
'                End If
'                Return penCount > 0
'            End If
'            Return False
'        Catch ex As Exception
'            'ModernDialog.ShowMessage(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButton.OK);
'            Throw
'            Return False
'        End Try
'    End Function

'    Private Shared Sub AddLog(sLog As String)
'        If (Not baseLogWriter Is Nothing) Then
'            baseLogWriter.WriteLine(sLog)
'        End If
'    End Sub

'    Public Shared Sub StartLog(sKey As String)
'        Dim sFile As String = Path.Combine(baseLogPath, sKey + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt")

'        StopLog()
'        baseLogWriter = New StreamWriter(sFile, False, Encoding.UTF8)
'    End Sub

'    Public Shared Sub StartLog(programInfo As TargetInfo)
'        Dim sKey As String = programInfo.ProgramCode + "_" + programInfo.ProgramNameKanji + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"
'        Dim sFile As String = Path.Combine(baseLogPath, sKey)

'        StopLog()
'        baseLogWriter = New StreamWriter(sFile, False, Encoding.UTF8)
'    End Sub
'    Public Shared Sub StopLog()
'        If (Not baseLogWriter Is Nothing) Then
'            GetPendingChange()

'            baseLogWriter.Close()
'            baseLogWriter = Nothing
'        End If
'    End Sub

'    Public Shared Sub GetPendingChange()
'        Dim pendingChanges As PendingChange() = baseWorkspace.GetPendingChanges()
'        If (Not (pendingChanges Is Nothing) AndAlso pendingChanges.Length > 0) Then
'            For Each pendingChange As PendingChange In pendingChanges
'                AddLog("Path: " + pendingChange.LocalItem + ", PendingChange: " + pendingChange.GetLocalizedStringForChangeType(pendingChange.ChangeType))
'            Next
'        End If
'    End Sub
'End Class
