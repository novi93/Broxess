Imports System.Data
Imports System.IO
Imports System.Text

Public Class FileHelper
    Protected Shared ENCODE As Encoding = Encoding.UTF8

    Shared Sub Output(ouputData As DataEntity, outputInfo As OutputInfo)
        Try
            ''todo check 
            With ouputData.MvcData
                Dim xFolder = Path.Combine(outputInfo.Path, outputInfo.MvcFodler)
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                Dim xFileName = String.Format("MVC.sql")
                Dim xPath = Path.Combine(xFolder, xFileName)
                Save2File(ouputData.MvcData, xPath)
            End With

            With ouputData.ProxyData
                Dim xFolder = Path.Combine(outputInfo.Path, outputInfo.ProxyFolder)
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                Dim xFileName = String.Format("Proxy.sql")
                Dim xPath = Path.Combine(xFolder, xFileName)
                Save2File(ouputData.ProxyData, xPath)
            End With

            With ouputData.ProcessData
                Dim xFolder = Path.Combine(outputInfo.Path, outputInfo.ProcessFolder)
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                Dim xFileName = String.Format("Process.sql")
                Dim xPath = Path.Combine(xFolder, xFileName)
                Save2File(ouputData.ProcessData, xPath)
            End With
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Shared Sub OutputToiawase(ouputData As DataEntity, outputInfo As OutputInfo, Optional isOnefile As Boolean = False)
        Try
            With ouputData.MvcData
                Dim xFolder = Path.Combine(outputInfo.Path, outputInfo.MvcFodler)
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                For Each dt As DataTable In .Tables
                    Dim xFileName = String.Format("{0}.sql", dt.TableName)
                    Dim xPath = Path.Combine(xFolder, xFileName)
                    Save2File_Dt(dt, xPath, isOnefile)
                Next
            End With

            With ouputData.ProxyData
                Dim xFolder = Path.Combine(outputInfo.Path, outputInfo.ProxyFolder)
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                For Each dt As DataTable In .Tables
                    Dim xFileName = String.Format("{0}.sql", dt.TableName)
                    Dim xPath = Path.Combine(xFolder, xFileName)
                    Save2File_Dt(dt, xPath, isOnefile)
                Next
            End With

            With ouputData.ProcessData
                Dim xFolder = Path.Combine(outputInfo.Path, outputInfo.ProcessFolder)
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                For Each dt As DataTable In .Tables
                    Dim xFileName = String.Format("{0}.sql", dt.TableName)
                    Dim xPath = Path.Combine(xFolder, xFileName)
                    Save2File_Dt(dt, xPath, isOnefile)
                Next
            End With
        Catch ex As Exception
            Throw
        End Try
    End Sub


    Shared Sub Output(ds As DataSet, path As String)
        Try
            ''todo check 
            With ds
                Dim xFolder = path
                If Not Directory.Exists(xFolder) Then
                    Directory.CreateDirectory(xFolder)
                End If
                Dim xFileName = String.Format("Output.sql")
                Dim xPath = System.IO.Path.Combine(xFolder, xFileName)
                Save2File(ds, xPath)
            End With

        Catch ex As Exception
            Throw
        End Try
    End Sub
    Shared Function Output(dt As DataTable, path As String) As Boolean
        Try
            If dt.Rows.Count <= 0 Then
                Return False
            End If
            Dim xFileName = String.Format("dbo.{0}.sql", dt.TableName)
            Dim xPath = System.IO.Path.Combine(path, xFileName)
            Save2File_Dt(dt, xPath)
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "CSV出力"

    Public Shared Sub Save2File(ByVal ds As DataSet, pPath As String)
        Try
            If ds.Tables.Count = 0 Then
                Return
            End If
            ''CSV形式ファイル作成
            Dim lineList As New ArrayList
            Dim xLine As String

            ''文字列初期化
            xLine = String.Empty

            For Each dR As DataRow In ds.Tables("DELETE_ALL").Rows

                ''文字列初期化
                xLine = dR("SQLSTRING")
                ''文字列を配列にセット
                lineList.Add(String.Format("{0}", xLine))
            Next

            For Each dt As DataTable In ds.Tables
                If dt.TableName <> "DELETE_ALL" Then
                    For Each dR As DataRow In dt.Rows

                        ''文字列初期化
                        xLine = dR("SQLSTRING")
                        ''文字列を配列にセット
                        lineList.Add(String.Format("{0}", xLine))
                    Next
                End If
            Next

            ''テキストファイル作成
            File.WriteAllLines(pPath, DirectCast(lineList.ToArray(GetType(String)), String()), ENCODE)

        Catch ex As Exception
            Throw
        End Try

    End Sub

    Public Shared Sub Save2File_Dt(ByVal dt As DataTable, pPath As String, Optional isAppend As Boolean = False)
        Try
            If dt.Rows.Count <= 0 Then
                Return
            End If
            ''CSV形式ファイル作成
            Dim lineList As New ArrayList
            Dim xLine As String

            ''文字列初期化
            xLine = String.Empty

            For Each dR As DataRow In dt.Rows
                ''文字列初期化
                xLine = dR("SQLSTRING")
                ''文字列を配列にセット
                lineList.Add(String.Format("{0}", xLine))
            Next
            If isAppend Then
                File.AppendAllLines(pPath, DirectCast(lineList.ToArray(GetType(String)), String()), ENCODE)
            Else
                ''テキストファイル作成
                File.WriteAllLines(pPath, DirectCast(lineList.ToArray(GetType(String)), String()), ENCODE)
            End If

        Catch ex As Exception
            Throw
        End Try

    End Sub
#End Region
End Class
