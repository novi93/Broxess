Imports System.Data
Imports SiU.Process.Common.Functions.CommonMethod
Imports SiU.Process.Common.Functions.SystemValue
Imports System.Text
Imports System.IO

Public Class SqlGenerator

#Region "メソッド"

    Protected Shared Function GetSqlDeleteDataSet(dsMaster As DataSet) As DataTable

        If dsMaster Is Nothing Then
            Throw New NullReferenceException()
        End If
        Dim nUpdSyubetu As Integer = 1
        'Dim dsMaster As New DataSet
        'dsMaster.Tables.Add(pDt.Copy)

        Dim nGno As Integer = 0
        Dim dt As New DataTable
        Dim xSqlTxt As String = String.Empty
        Dim xSqlCol As String = String.Empty
        Dim xSqlValue As String = String.Empty
        'Dim MyMainTable As String = pDt.TableName

        'テーブル定義
        dt.Columns.Add("GNO", GetType(Integer))
        dt.Columns.Add("SQLSTRING", GetType(String))
        dt.Columns.Add("UPD_TABLE_NAME", GetType(String))
        dt.Columns.Add("GYO_KKBN", GetType(Integer))
        dt.Columns.Add("UPD_KEYSTRING", GetType(String))
        dt.Columns.Add("UPD_COLSTRING", GetType(String))
        dt.TableName = "DELETE_ALL"

        If dsMaster.Tables.Count > 0 Then
            For index = dsMaster.Tables.Count - 1 To 0 Step -1
                Call GetDeleteSqlDataSet(dsMaster.Tables(index), dt)
            Next
        End If

        Return dt

    End Function

    ''' <summary>
    ''' Sql文のデータセット作成
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>更新SQLのデータセット作成</remarks>
    Protected Shared Function GetSqlDataSet(pDt As DataTable) As DataTable

        If pDt Is Nothing Then
            Throw New NullReferenceException()
        End If
        Dim nUpdSyubetu As Integer = 1
        Dim dsMaster As New DataSet
        dsMaster.Tables.Add(pDt.Copy)

        Dim nGno As Integer = 0
        Dim dt As New DataTable
        Dim xSqlTxt As String = String.Empty
        Dim xSqlCol As String = String.Empty
        Dim xSqlValue As String = String.Empty
        Dim MyMainTable As String = pDt.TableName

        'テーブル定義
        dt.Columns.Add("GNO", GetType(Integer))
        dt.Columns.Add("SQLSTRING", GetType(String))
        dt.Columns.Add("UPD_TABLE_NAME", GetType(String))
        dt.Columns.Add("GYO_KKBN", GetType(Integer))
        dt.Columns.Add("UPD_KEYSTRING", GetType(String))
        dt.Columns.Add("UPD_COLSTRING", GetType(String))
        dt.TableName = pDt.TableName
        Try

            '-------------------------------------------------------------------
            '新規/修正時はINSERT文を生成
            If nUpdSyubetu <> 3 Then
                Try

                    'テーブル毎
                    For i As Integer = 0 To dsMaster.Tables.Count - 1
                        'Call GetDeleteSqlDataSet(pDt, dt)

                        '※※それぞれ合ったINSERT文を生成
                        xSqlTxt = String.Format("INSERT INTO [dbo].[{0}]", MyMainTable)
                        'データテーブルへ行追加
                        Dim dr As DataRow

                        '行毎
                        For j As Integer = 0 To dsMaster.Tables(i).Rows.Count - 1
                            dr = dt.NewRow
                            '行番号加算
                            nGno = nGno + 1
                            '初期化
                            xSqlCol = " ("
                            xSqlValue = "("
                            '列毎
                            For k As Integer = 0 To dsMaster.Tables(i).Columns.Count - 1
                                '列名
                                xSqlCol = xSqlCol & String.Format("[{0}]", fToText(dsMaster.Tables(i).Columns(k).ColumnName))

                                '値
                                Select Case fToText(dsMaster.Tables(i).Columns(k).ColumnName)
                                    'Case "UD_DATE"
                                    '    xSqlValue = xSqlValue & "GETDATE()"
                                    'Case "AD_DATE"  ''新規時はSYSDATE、修正時は前回の値
                                    '    If nUpdSyubetu = 1 Then
                                    '        xSqlValue = xSqlValue & "GETDATE()"
                                    '    Else
                                    '        xSqlValue = xSqlValue & fSqlStr(fToDate(dsMaster.Tables(i).Rows(j).Item(k)), BrowsefSqlStr.fSqlStr_DateTime)
                                    '    End If
                                    'Case "OPID"
                                    '    xSqlValue = xSqlValue & fSqlStr(SysOpid, BrowsefSqlStr.fSqlStr_Text)
                                    'Case "UD_USER"
                                    '    xSqlValue = xSqlValue & fSqlStr(SysUserNo, BrowsefSqlStr.fSqlStr_Long)
                                    'Case "YOBI_DATE1", "YOBI_DATE2", "YOBI_DATE3", "YOBI_DATE4", "YOBI_DATE5", _
                                    '     "YOBI_DATE6", "YOBI_DATE7", "YOBI_DATE8", "YOBI_DATE9", "YOBI_DATE10"
                                    '    If fToDate(dsMaster.Tables(i).Rows(j).Item(k)) = SysMinDate Then
                                    '        xSqlValue = xSqlValue & "NULL"
                                    '    Else
                                    '        xSqlValue = xSqlValue & fSqlStr(fToDate(dsMaster.Tables(i).Rows(j).Item(k)), BrowsefSqlStr.fSqlStr_DateTime)
                                    '    End If
                                    Case Else
                                        If IsDBNull(dsMaster.Tables(i).Rows(j).Item(k)) Then
                                            'NULL時はNULLで更新
                                            xSqlValue = xSqlValue & "NULL"
                                        Else
                                            Select Case dsMaster.Tables(i).Columns(k).DataType.Name
                                                Case "Int16", "Int32", "Int64", "Integer", "Double", "Decimal"  ''数値
                                                    xSqlValue = xSqlValue & fSqlStr(fToNumber(dsMaster.Tables(i).Rows(j).Item(k)), BrowsefSqlStr.fSqlStr_Decimal)
                                                Case "DateTime"                                                 ''日付
                                                    'xSqlValue = xSqlValue & fSqlStr(fToDate(dsMaster.Tables(i).Rows(j).Item(k)), BrowsefSqlStr.fSqlStr_DateTime)
                                                    xSqlValue = xSqlValue & String.Format("N'{0}'", SiuFormat(dsMaster.Tables(i).Rows(j).Item(k), "yyyy-MM-dd HH:mm:ss"))
                                                Case "Guid"                                                 ''Guid
                                                    xSqlValue = xSqlValue & String.Format("N'{0}'", dsMaster.Tables(i).Rows(j).Item(k))
                                                Case "Boolean", "Bit"                                                 ''Guid
                                                    xSqlValue = xSqlValue & String.Format("{0}", If(dsMaster.Tables(i).Rows(j).Item(k), 1, 0))
                                                Case Else   ''文字列など
                                                    If dsMaster.Tables(i).Columns(k).ColumnName = "CODE" Then
                                                        xSqlValue = xSqlValue & fSqlStr(fCodeformat(dsMaster.Tables(i).Rows(j).Item(k), SysCharKeta), BrowsefSqlStr.fSqlStr_Text)
                                                    Else
                                                        xSqlValue = xSqlValue & fSqlStr(fToText(dsMaster.Tables(i).Rows(j).Item(k)), BrowsefSqlStr.fSqlStr_Text)
                                                    End If
                                            End Select
                                        End If
                                End Select

                                'カンマ・括弧
                                If k = dsMaster.Tables(i).Columns.Count - 1 Then    ''最終列
                                    xSqlCol = xSqlCol & ") VALUES "
                                    xSqlValue = xSqlValue & ")"
                                Else
                                    xSqlCol = xSqlCol & ", "
                                    xSqlValue = xSqlValue & ", "
                                End If
                            Next
                            'INSERT文合成

                            'データテーブルへ格納
                            dr.Item(0) = nGno
                            dr.Item(1) = xSqlTxt & xSqlCol & xSqlValue
                            dr.Item(2) = MyMainTable
                            dr.Item(3) = 1      '追加
                            dt.Rows.Add(dr)
                        Next
                    Next
                Catch ex As Exception
                    Throw New Exception("MakeInsTxt_Err")
                End Try
            End If

            '-------------------------------------------------------------------

        Catch ex As Exception
            Select Case ex.Message
                Case "MakeDeleteTxt_Err"
                Case "MakeInsTxt_Err"
            End Select
            dt = Nothing
            Return dt
        End Try

        Return dt

    End Function

    Public Shared Function GetColString(ByVal xTableName As String) As String
        Dim xCol As String = String.Empty

        Select Case xTableName
            'Case "Programs"
            '    xCol = "Id"
            'Case "ToiawaseColumnProperties"
            '    xCol = "ProgramId,UserId,Code,No"
            'Case "ToiawaseDefinations"
            '    xCol = "ProgramId"
            'Case "ToiawaseProperties"
            '    xCol = "Id"
            'Case "ToiawaseReportProperties"
            '    xCol = "ProgramId,UserId,Code"
            'Case "ToiawaseSearchPanels"
            '    xCol = "ProgramId,ParameterNo"
            Case "Programs"
                xCol = "Id"
            Case "ToiawaseColumnProperties"
                xCol = "ProgramId"
            Case "ToiawaseDefinations"
                xCol = "ProgramId"
            Case "ToiawaseProperties"
                xCol = "ProgramId"
            Case "ToiawaseReportProperties"
                xCol = "ProgramId"
            Case "ToiawaseSearchPanels"
                xCol = "ProgramId"
            Case "M_PROGRAM_EXTENSION"
                xCol = "Id"
            Case "T_MENU_EXTENSION"
                xCol = "PROGRAM_CODE"
            Case "Menu"
                xCol = "ProgramId"
            Case "M_PROGRAM_ALLOW"
                xCol = "PROGRAM"
            Case "T_PROGRAM_LINK"
                xCol = "CODE"
            Case Else
                Throw New ArgumentException("at GetColString, Table not define colstring ", xTableName)
        End Select

        Return xCol

    End Function

    Public Shared Function GetKeyString(ByVal pCols As String, ByVal pDr As DataRow) As String
        Dim xKeys As String = String.Empty

        For Each xCol As String In pCols.Split(","c)
            'キー文字列生成
            If TypeOf pDr.Item(xCol) Is String Then
                xKeys &= fSqlStr(fToText(pDr.Item(xCol)), BrowsefSqlStr.fSqlStr_Text) & ","
            ElseIf TypeOf pDr.Item(xCol) Is Guid Then
                xKeys &= String.Format("N'{0}'", pDr.Item(xCol))
                'xKeys &= fSqlStr(pDr.Item(xCol).ToString(), BrowsefSqlStr.fSqlStr_Text) & ","
            Else
                xKeys &= fSqlStr(pDr.Item(xCol), BrowsefSqlStr.fSqlStr_Integer) & ","
            End If
        Next
        xKeys = xKeys.Substring(0, xKeys.Length - 1)

        Return xKeys

    End Function

    Public Shared Sub GetDeleteSqlDataSet(ByVal pDt As DataTable, ByRef dt As DataTable)
        Dim nGno As Integer = 0
        Dim xSqlTxt As String = String.Empty
        Dim dRUw As DataRow = Nothing
        Dim xKeyStringLocal As String = String.Empty
        Dim xCols As String = String.Empty
        Dim dR As DataRow = Nothing

        If pDt.Rows.Count > 0 Then
            xCols = GetColString(pDt.TableName)
            'For j As Integer = 0 To pDt.Rows.Count - 1
            ''行番号加算
            nGno += 1
            ''データテーブルへ行定義追加
            dR = dt.NewRow
            ''キー文字列の作成
            xKeyStringLocal = GetKeyString(xCols, pDt.Rows(0))
            ''※※それぞれ合ったDELETE文を生成
            xSqlTxt = CreateDeleteDataSQL(fToText(pDt.TableName), xKeyStringLocal, xCols)
            ''データテーブルへ格納
            dR.Item(0) = nGno
            dR.Item(1) = xSqlTxt
            dR.Item(2) = fToText(pDt.TableName)
            dR.Item(3) = 3      '削除
            dR.Item(4) = xKeyStringLocal      'キー文字列
            dR.Item(5) = xCols      '項目文字列
            dt.Rows.Add(dR)

            'Next
        End If

    End Sub

    Public Shared Function CreateDeleteDataSQL(ByVal TableName As String, ByVal xKey As String, ByVal xCol As String) As String

        Dim xSqlTxt As String = String.Empty

        xSqlTxt = xSqlTxt & vbCrLf & "DELETE FROM dbo." & TableName
        xSqlTxt = xSqlTxt & GetSqlWhere(TableName, xKey, xCol)

        Return xSqlTxt

    End Function

    Public Shared Function GetSqlWhere(ByVal pTableName As String, ByVal pKeys As String, ByVal pCols As String) As String
        Dim xSqlTxt As String = String.Empty
        Dim pKeysArry() As String = Nothing
        Dim pColsArry() As String = Nothing

        pKeysArry = pKeys.Split(","c)
        pColsArry = pCols.Split(","c)

        xSqlTxt = ""
        xSqlTxt = xSqlTxt & " WHERE 1 = 1"
        For j As Integer = 0 To pColsArry.Count - 1
            xSqlTxt = xSqlTxt & " AND "
            xSqlTxt = xSqlTxt & String.Format("[{0}]", fToText(pColsArry(j)))
            xSqlTxt = xSqlTxt & " = " & pKeysArry(j)
        Next

        Return xSqlTxt
    End Function

    Public Shared Function GenerateScript(dt As DataTable) As DataTable
        Try
            'SQL文作成
            Return GetSqlDataSet(dt)
        Catch ex As Exception
            Throw
        End Try

    End Function
    Public Shared Function GenerateScriptDs(ds As DataSet) As DataSet
        Try
            Dim rs As New DataSet
            Dim rsDeleteDt = GetSqlDeleteDataSet(ds).Copy
            rs.Tables.Add(rsDeleteDt)

            For Each dt As DataTable In ds.Tables
                Dim rsDt = GenerateScript(dt).Copy
                rsDt.TableName = dt.TableName
                rs.Tables.Add(rsDt)

            Next
            Return rs
        Catch ex As Exception
            Throw
        End Try

    End Function
#End Region

End Class
