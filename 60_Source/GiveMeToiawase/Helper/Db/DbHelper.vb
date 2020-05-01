Imports System.Data
Imports System.Data.SqlClient
Imports SiU.Process.Common.Functions.CommonMethod
Imports SiU.Process.Common.Functions.SystemValue
Imports System.Data.Common
Imports System.IO
Imports System.Transactions

Public Class DbHelper
    Protected Shared MvcRepository As Repository
    Protected Shared ProxyRepository As Repository
    Protected Shared ProcessRepository As Repository

    Public Shared Sub Init()
        MvcRepository = New Repository(New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("UsoliaMvc").ConnectionString))
        ProxyRepository = New Repository(New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ProxyConnection").ConnectionString))
        ProcessRepository = New Repository(New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ProcessConnection").ConnectionString))
    End Sub

    Shared Sub Init(connectInfo As ConnectionInfo)
        MvcRepository = New Repository(New SqlConnection(connectInfo.mvc))
        ProxyRepository = New Repository(New SqlConnection(connectInfo.proxy))
        ProcessRepository = New Repository(New SqlConnection(connectInfo.process))

    End Sub

    Shared Function GetCorrectTableName(tableName As String, repo As Repository) As String
        Dim xSql = String.Format(" SELECT TABLE_NAME " &
                                 " FROM INFORMATION_SCHEMA.TABLES " &
                                 " WHERE TABLE_TYPE = 'BASE TABLE'" &
                                 " AND UPPER(TABLE_NAME) = UPPER('{0}')",
                                 tableName)
        Dim dt = repo.GetData(xSql)
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("TABLE_NAME")
        Else
            Throw New Exception(String.Format("Table [{0}] not exist in Database",
                                              tableName))
        End If
    End Function
#Region "Public"
    Public Shared Function LoadData(request As RequestInfo) As DataEntity
        Dim rs As New DataEntity
        With rs
            .MvcData = LoadMvc(request)
            .ProxyData = LoadProxy(request)
            .ProcessData = LoadProcess(request)
        End With
        Return rs
    End Function

    Shared Sub UpdateDB(outputInfo As OutputInfo)
        Dim mvcTran = MvcRepository.BeginTrans()
        Dim proxyTran = ProxyRepository.BeginTrans()
        Dim processTran = ProcessRepository.BeginTrans()
        Try
            With outputInfo
                MvcRepository.excutedFolder(Path.Combine(.Path, .MvcFodler), mvcTran)
                ProxyRepository.excutedFolder(Path.Combine(.Path, .ProxyFolder), proxyTran)
                ProcessRepository.excutedFolder(Path.Combine(.Path, .ProcessFolder), processTran)
                MvcRepository.CommitTrans(mvcTran)
                ProxyRepository.CommitTrans(proxyTran)
                ProcessRepository.CommitTrans(processTran)
            End With
        Catch ex As Exception
            MvcRepository.RollBackTrans(mvcTran)
            ProxyRepository.RollBackTrans(proxyTran)
            ProcessRepository.RollBackTrans(processTran)
            Throw
        End Try
    End Sub

#End Region

#Region "Load"

#Region "MVC"
    Private Shared Function LoadMvc(request As RequestInfo) As DataSet
        Dim rs As New DataSet
        rs.Tables.Add(LoadPrograms(request))
        rs.Tables.Add(LoadToiawaseSearchPanels(request))
        rs.Tables.Add(LoadToiawaseReportProperties(request))
        rs.Tables.Add(LoadToiawaseProperties(request))
        rs.Tables.Add(LoadToiawaseColumnProperties(request))
        rs.Tables.Add(LoadToiawaseDefinations(request))
        rs.Tables.Add(LoadMenu(request))
        Return rs
    End Function

    Private Shared Function LoadPrograms(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.Programs
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [Id] = {1} ",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadToiawaseSearchPanels(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.ToiawaseSearchPanels
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [ProgramId] = {1} ORDER BY ParameterNo ",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadToiawaseReportProperties(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.ToiawaseReportProperties
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [ProgramId] = {1} ORDER BY UserId,Code ",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadToiawaseProperties(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.ToiawaseProperties
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [ProgramId] = {1} ORDER BY UserId,Code",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadToiawaseColumnProperties(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.ToiawaseColumnProperties
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [ProgramId] = {1} ORDER BY UserId,Code,No ",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadToiawaseDefinations(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.ToiawaseDefinations
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [ProgramId] = {1} ",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadMenu(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.Menu

        Dim xSql = String.Format("SELECT * FROM {0} WHERE [ProgramId] = {1} ORDER BY Id",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = MvcRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function


#End Region

#Region "Proxy"
    Private Shared Function LoadProxy(request As RequestInfo) As DataSet
        Dim rs As New DataSet
        rs.Tables.Add(LoadM_PROGRAM_EXTENSION(request))
        rs.Tables.Add(LoadT_MENU_EXTENSION(request))
        rs.Tables.Add(LoadT_PROGRAM_LINK(request))
        Return rs
    End Function
    Private Shared Function LoadM_PROGRAM_EXTENSION(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.M_PROGRAM_EXTENSION
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [Id] = {1} ORDER BY Id",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = ProxyRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadT_MENU_EXTENSION(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.T_MENU_EXTENSION
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [PROGRAM_CODE] = {1}",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = ProxyRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function
    Public Shared Function GetNewMenuId(parentId As String, programId As String) As DataTable
        Dim xTableName As String = TableName.Menu
        Dim xSql As String = String.Empty
        xSql = "SELECT MAX(MN.OrderNumber) AS OrderNumber, MAX(EX2.MENU_A) AS MENU_A, MAX(EX2.MENU_B) AS MENU_B, MAX(EX2.MENU_C) AS MENU_C FROM ProcessProxy.[dbo].[T_MENU_EXTENSION] EX1"
        xSql = xSql & vbCrLf & " LEFT JOIN [dbo].[T_MENU_EXTENSION] AS EX2"
        xSql = xSql & vbCrLf & " ON EX1.MENU_A = EX2.MENU_A"
        xSql = xSql & vbCrLf & " AND EX1.MENU_B = EX2.MENU_B"
        xSql = xSql & vbCrLf & " LEFT JOIN Menu AS MN"
        xSql = xSql & vbCrLf & " ON MN.Id = EX2.Id"
        xSql = xSql & vbCrLf & " WHERE EX1.Id = {0}"
        xSql = xSql & vbCrLf & " AND MN.ProgramId <> {1}"
        xSql = String.Format(xSql, fSqlStr(parentId, BrowsefSqlStr.fSqlStr_Text), fSqlStr(programId, BrowsefSqlStr.fSqlStr_Text))
        Dim dt = ProxyRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Public Shared Function GetNewMenuId(menuA As Integer, menuB As Integer, programId As String) As DataTable
        Dim xTableName As String = TableName.Menu
        Dim xSql As String = String.Empty
        xSql = "SELECT MAX(MENU_A) AS MENU_A, MAX(MENU_B) AS MENU_B ,MAX(MENU_C) AS MENU_C FROM [dbo].[T_MENU_EXTENSION]"
        xSql = xSql & vbCrLf & " WHERE MENU_A = {0}"
        xSql = xSql & vbCrLf & " AND MENU_B = {1}"
        xSql = xSql & vbCrLf & " AND PROGRAM_CODE <> {2}"
        xSql = String.Format(xSql, fSqlStr(menuA, BrowsefSqlStr.fSqlStr_Integer), fSqlStr(menuB, BrowsefSqlStr.fSqlStr_Integer), fSqlStr(programId, BrowsefSqlStr.fSqlStr_Text))
        Dim dt = ProxyRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function

    Private Shared Function LoadT_PROGRAM_LINK(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.T_PROGRAM_LINK
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [CODE] = {1} ORDER BY CODE,KKBN,PARAMETER",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = ProcessRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function
#End Region

#Region "Process"
    Private Shared Function LoadProcess(request As RequestInfo) As DataSet
        Dim rs As New DataSet
        rs.Tables.Add(LoadM_PROGRAM_ALLOW(request))
        Return rs
    End Function
    Private Shared Function LoadM_PROGRAM_ALLOW(request As RequestInfo) As DataTable
        Dim xTableName As String = TableName.M_PROGRAM_ALLOW
        Dim xSql = String.Format("SELECT * FROM {0} WHERE [PROGRAM] = {1} ORDER BY PROGRAM,USER_NO,GROUP_CODE,KENGEN",
                                 xTableName,
                                 fSqlStr(request.ProgramCode_Source))
        Dim dt = ProcessRepository.GetData(xSql)
        dt.TableName = xTableName
        Return dt
    End Function
#End Region

#End Region



End Class
