Option Strict Off
Imports System.Text
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Web.Services.Protocols
Imports System.Deployment.Application.ApplicationDeployment

Imports SiU.Process.Common.Functions.SystemValue
''2008.10.01 ADD START YONEI
Imports System.ComponentModel
'Imports SiU.Process.Common.Functions.Applications
Imports System.Deployment.Application
''2008.10.01 ADD END

''' <summary>
''' CommonMethod -- 共通関数
''' </summary>
''' <remarks></remarks>
''' <history version="5.0.0.0" date="" name="">初版不明</history>
''' <history version="5.0.0.0" date="2009.03.02" name="大内 登志也">部門コードの文字コード化対応</history>
''' <history version="5.0.0.0" date="2009.03.25" name="千葉 友則">関数の引数がNothingの場合、クラスの初期値（MinValue）として扱うように修正</history>
''' <history version="5.0.0.0" date="2009.06.01" name="高嶋 直樹">切り抜いた文字列のバイト数取得不正を修正</history>
''' <history version="5.0.0.0" date="2009.08.13" name="内田 久義">年なしの日付書式にもカルチャを指定</history>
''' <history version="5.0.0.0" date="2010.03.04" name="米井 優顕">fSqlStrTypeを追加</history>
''' <history version="5.0.0.0" date="2010.03.30" name="清水 勝也">製品名定数を追加</history>
''' <history version="5.0.0.0" date="2010.04.08" name="米井 優顕">ZeiCalculateを追加</history>
''' <history version="5.0.0.0" date="2010.05.13" name="清水 勝也">部門権限に関する共通メソッドを追加</history>
''' <history version="5.0.0.0" date="2010.05.26" name="清水 勝也">(課題表:3091)SiuFormatの和暦の扱いでエラーが無限ループするのを修正</history>
''' <history version="5.0.0.0" date="2010.07.02" name="清水 勝也">fSqlStrをoverloadしてDataColumnから型・PROCESSでの用途を判断して置換する処理を追加
'''                                                               DataRowからINSERT/UPDATE/DELETEを生成するfSqlStrInsert等を追加</history>
''' <history version="5.0.0.0" date="2010.07.26" name="HANH-LV">現場システムオフライン対応</history>
''' <history version="5.0.1.1" date="2011.12.14" name="内田 久義">(SCで落ちる原因になっていた為)EventLogへの書き込みエラーは何もしようがないのでスルーする</history>
''' <history version="5.0.1.1" date="2014.03.12" name="内田 久義">時刻の書式指定が不正な為、12時間制でフォーマットされる(参照問題が発生するかもしれないのでバージョンは上げません)</history>
''' <history version="5.0.1.1" date="2014.06.26" name="TRI-PQ">(TID.3220-IID.26713)SSL化(参照問題が発生するかもしれないのでバージョンは上げません)</history>
''' <history version="6.0.0.0" date="2017.08.24" name="HAI-NM">SQLをPROCES.S5からSERVER2016(プロジェクト武蔵)に移行する。</history>
Public Class CommonMethod

    ''2010.03.30 ADD START SHIMIZU
    ''' <summary>
    ''' ProductName -- 製品名
    ''' </summary>
    ''' <remarks>変更禁止</remarks>
    Public Const ProductName As String = "Ussol.Process"
    ''2010.03.30 ADD END

    '' 2010.07.26 ADD START HANH-LV
    '' 対応内容：現場システムオフライン対応
    ''' <summary>オフラインモードの識別</summary>
    ''' <remarks> </remarks>
    Public Shared IsOfflineMode As Boolean = False
    '' 2010.07.26 ADD END HANH-LV

#Region " 列挙体 "
    ''***************************************************************************************************************************************************

    ''' <summary>BrowsefSqlStr -- ORACLE SQL列修飾用変数</summary>
    ''' <remarks></remarks>
    Public Enum BrowsefSqlStr
        ''' <summary>BigInt型</summary>
        fSqlStr_BigInt = 16
        ''' <summary>Binary型</summary>
        fSqlStr_Binary = 9
        ''' <summary>Boolean型</summary>
        fSqlStr_Boolean = 1
        ''' <summary>Byte型</summary>
        fSqlStr_Byte = 2
        ''' <summary>Char型</summary>
        fSqlStr_Char = 18
        ''' <summary>Currency型</summary>
        fSqlStr_Currency = 5
        ''' <summary>DATE型</summary>
        fSqlStr_DATE = 8
        ''' <summary>Decimal型</summary>
        fSqlStr_Decimal = 20
        ''' <summary>Double型</summary>
        fSqlStr_Double = 7
        ''' <summary>Float型</summary>
        fSqlStr_Float = 21
        ''' <summary>GUID型</summary>
        fSqlStr_GUID = 15
        ''' <summary>Integer型</summary>
        fSqlStr_Integer = 3
        ''' <summary>Long型</summary>
        fSqlStr_Long = 4
        ''' <summary>LongBinary型</summary>
        fSqlStr_LongBinary = 11
        ''' <summary>Memo型</summary>
        fSqlStr_Memo = 12
        ''' <summary>Numeric型</summary>
        fSqlStr_Numeric = 19
        ''' <summary>Single型</summary>
        fSqlStr_Single = 6
        ''' <summary>Text型</summary>
        fSqlStr_Text = 10
        ''' <summary>Time型</summary>
        fSqlStr_Time = 22
        ''' <summary>TimeStamp型</summary>
        fSqlStr_TimeStamp = 23
        ''' <summary>VarBinary型</summary>
        fSqlStr_VarBinary = 17
        ''' <summary>DateTime型</summary>
        fSqlStr_DateTime = 51
    End Enum

    ''' <summary>BrowsefToNumber -- fToNumber返却値の数値ﾀｲﾌﾟ</summary>
    ''' <remarks></remarks>
    Public Enum BrowsefToNumber
        ''' <summary>Integer型に変換</summary>
        fToNumber_INTEGER = 1
        ''' <summary>Long型に変換</summary>
        fToNumber_LONG = 2
        ''' <summary>Single型に変換</summary>
        fToNumber_SINGLE = 3
        ''' <summary>Double型に変換</summary>
        fToNumber_DOUBLE = 4
        ''' <summary>Decimal型に変換</summary>
        fToNumber_DECIMAL = 5
        ''' <summary>Short型に変換</summary>
        fToNumber_SHORT = 6
    End Enum

    ''' <summary>BrowsefToDate -- fToDate返却値の数値ﾀｲﾌﾟ</summary>
    ''' <remarks></remarks>
    Public Enum BrowsefToDate
        ''' <summary>日付時刻で返却</summary>
        fToDate_DateTime = 1
        ''' <summary>日付のみで返却</summary>
        fToDate_DateOnly = 2
        ''' <summary>時刻のみで返却</summary>
        fToDate_TimeOnly = 3
    End Enum

    ''' <summary>BrowseHasu -- 端数処理区分</summary>
    ''' <remarks></remarks>
    Public Enum BrowseHasu
        ''' <summary>切り上げ</summary>
        bHasu_Kiriage = 1
        ''' <summary>切捨て</summary>
        bHasu_Kirisute = 2
        ''' <summary>四捨五入</summary>
        bHasu_Shisyagonyu = 3
    End Enum

    ''' <summary>
    ''' ラベル状態区分
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum BrouseConfig
        ''' <summary>未登録文字</summary>
        NoData = 1
        ''' <summary>新規登録文字(入力開始前)</summary>
        NewData = 2
        ''' <summary>新規登録文字(入力開始後)</summary>
        NewDataInput = 3
        ''' <summary>修正登録文字(入力開始前)</summary>
        UpdData = 4
        ''' <summary>修正登録文字(入力開始後)</summary>
        UpdDataInput = 5
        ''' <summary>権限無し文字</summary>
        NoSec = 6
        ''' <summary>更新中文字</summary>
        Running = 7
        ''' <summary>初期化(ラベルの非表示)</summary>
        Initialize = 99
    End Enum

    ''' <summary>ControlKbn -- コントロール区分</summary>
    ''' <remarks>コントロールが単独または範囲（開始/終了）の判定フラグ</remarks>
    ''' <history version="1" date="2008.05.23" name="千葉 友則">新規作成</history>
    Public Enum ControlKbn As Integer
        ''' <summary>Only</summary>
        ''' <remarks>単独指定</remarks>
        Only = 0
        ''' <summary>From</summary>
        ''' <remarks>開始値指定</remarks>
        From = 1
        ''' <summary>To</summary>
        ''' <remarks>終了値指定</remarks>
        [To] = 2
    End Enum

    ''' <summary>SysCodeType -- システムコード区分</summary>
    ''' <remarks>各種システムコード（SysMin***/SysMax***）を識別</remarks>
    ''' <history version="1" date="2008.05.23" name="千葉 友則">新規作成</history>
    ''' <history version="2" date="2009.03.02" name="大内 登志也">部門コードの文字コード化対応</history>
    Public Enum SysCodeType As Integer
        ''' <summary>Code</summary>
        ''' <remarks>SysMinCode/SysMaxCode</remarks>
        Code = 0
        ''' <summary>Date</summary>
        ''' <remarks>SysMinDate/SysMaxDate</remarks>
        [Date] = 1
        ''' <summary>SKoj</summary>
        ''' <remarks>SysMinSKoj/SysMaxSKoj</remarks>
        SKoj = 2
        ''' <summary>Koj</summary>
        ''' <remarks>SysMinKoj/SysMaxKoj</remarks>
        Koj = 3
        ''' <summary>KojUw</summary>
        ''' <remarks>SysMinKojUw/SysMaxKojUw</remarks>
        KojUw = 4
        ''' <summary>Tor</summary>
        ''' <remarks>SysMinTor/SysMaxTor</remarks>
        Tor = 5
        ''' <summary>Soko</summary>
        ''' <remarks>SysMinSoko/SysMaxSoko</remarks>
        Soko = 6
        ''' <summary>Syohin</summary>
        ''' <remarks>SysMinSyohin/SysMaxSyohin</remarks>
        Syohin = 7
        ''' <summary>Jyu</summary>
        ''' <remarks>SysMinJyu/SysMaxJyu</remarks>
        Jyu = 8
        ''' <summary>Hac</summary>
        ''' <remarks>SysMinHac/SysMaxHac</remarks>
        Hac = 9
        ''' <summary>Syu</summary>
        ''' <remarks>SysMinSyu/SysMaxSyu</remarks>
        Syu = 10
        ''' <summary>Uri</summary>
        ''' <remarks>SysMinUri/SysMaxUri</remarks>
        Uri = 11
        ''' <summary>Nyu</summary>
        ''' <remarks>SysMinNyu/SysMaxNyu</remarks>
        Nyu = 12
        ''' <summary>Sir</summary>
        ''' <remarks>SysMinSir/SysMaxSir</remarks>
        Sir = 13
        ''' <summary>Bmn</summary>
        ''' <remarks>SysMinBmn/SysMaxBmn</remarks>
        Bmn = 14
    End Enum

    ''' <summary>
    ''' UpdType -- 更新区分(1:新規, 2:更新, 3:削除)
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="清水 勝也">新規作成</history>
    Public Enum UpdType As Integer
        Insert = 1
        Update = 2
        Delete = 3
    End Enum

    ''' <summary>ActionType</summary>
    ''' <remarks>アクション定数</remarks>
    ''' <history version="1" date="2008.06.20" name="千葉 友則">新規作成</history>
    Public Enum ActionType As Integer
        DrillDown = 0
        Preview = 6
        Print = 7
        ExportPDF = 8
        ExportXLS = 9

        F1 = 1
        F2 = 2
        F3 = 3
        F4 = 4
        F5 = 5
        F6 = 6
        F7 = 7
        F8 = 8
        F9 = 9
        F10 = 10
        F11 = 11
        F12 = 12

        '' Funcキー未対応のｱｸｼｮﾝ定数は100～
        ExportHTML = 101

    End Enum

    ''2010.05.13 ADD START SHIMIZU 部門権限に関する共通メソッドの引数をBooleanではなくInteger(列挙体使用)に修正
    ''' <summary>
    ''' CheckBmnSecType -- 部門権限に関する共通メソッドの引数
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history version="2" date="2010.05.13" name="清水 勝也">新規作成</history>
    Public Enum CheckBmnSecType As Integer
        Sansyo = 1
        Input = 2
    End Enum
    ''2010.05.13 ADD END

#End Region

#Region " 置換関数 "
    ''***************************************************************************************************************************************************

    ''' <summary>fReplaceString -- 文字列の中の文字列を置換します。</summary>
    ''' <param name="pString">処理する文字列</param>
    ''' <param name="pShString">検索する文字列</param>
    ''' <param name="pRepString">置換する文字列</param>
    ''' <param name="fCaseSensitive">大文字／小文字を区別  する場合はTrue、区別しない場合はFalse</param>
    ''' <returns>置換後の文字列</returns>
    ''' <remarks>文字列の中の文字列を置換します。</remarks>
    ''' <作成者>Total VB SourceBook 6</作成者>
    ''' <備考></備考>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    <DebuggerStepThrough()>
    Public Shared Function fReplaceString(ByVal pString As String, ByVal pShString As String, ByVal pRepString As String, Optional ByVal fCaseSensitive As Boolean = True) As String
        Dim strTmp As String
        Dim intPos As Integer
        Dim StringComparison As System.StringComparison

        If pString Is Nothing Then
            pString = String.Empty
        End If
        If pShString Is Nothing Then
            pShString = String.Empty
        End If
        If pRepString Is Nothing Then
            pRepString = String.Empty
        End If

        '' 与えられたパラメータから比較タイプを設定
        If fCaseSensitive Then
            'StringComparison = System.StringComparison.CurrentCulture

            '大文字/小文字の区別をする場合はString.Replaceを使用する
            Return pString.Replace(pShString, pRepString)
        Else
            StringComparison = System.StringComparison.CurrentCultureIgnoreCase
        End If
        strTmp = pString

        Try

            '' 最初の検索を行います。
            intPos = strTmp.IndexOf(pShString, 0, StringComparison)

            If pRepString.Length = 0 Then
                strTmp = pString.Replace(pShString, pRepString)
            Else
                Do While intPos > -1
                    ' 文字列の検索を繰り返します。
                    strTmp = strTmp.Substring(0, intPos) & pRepString & strTmp.Substring(intPos + pRepString.Length, strTmp.Length - (intPos + pRepString.Length))
                    ' 次の検索を行います。
                    intPos = strTmp.IndexOf(pShString, intPos + pRepString.Length, StringComparison)
                Loop
            End If


        Catch ex As Exception

            Throw New Exception("fReplaceString:文字列置換に失敗しました｡" & vbCrLf & ex.Message)
            strTmp = pString

        End Try

        '' 値を返します。
        Return strTmp

    End Function

    ''' <summary>fReplaceString -- 文字列の中の文字列を置換します。</summary>
    ''' <param name="pString">処理する文字列</param>
    ''' <param name="pShString">検索する文字列</param>
    ''' <param name="pRepString">置換する文字列</param>
    ''' <param name="pOption">大文字／小文字を区別  する場合は0、区別しない場合は1</param>
    ''' <returns>置換後の文字列</returns>
    ''' <remarks>
    ''' 文字列の中の文字列を置換します。
    ''' ※fChangeStrの名称を変更（同じ機能は同じ名前に収束）
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Function fReplaceString(ByVal pString As String, ByVal pShString As String, ByVal pRepString As String, ByVal pOption As Integer) As String

        Return fReplaceString(pString, pShString, pRepString, (pOption = 0))

    End Function

    ''' <summary>fSqlStr -- オラクルSQL文字列化</summary>
    ''' <param name="pSourceString">元変数</param>
    ''' <param name="pSourceStringType">変数タイプ</param>
    ''' <returns>オラクルSQLに組み込める文字列</returns>
    ''' <remarks>データタイプ別にオラクルsql文字列を作成して戻します</remarks>
    ''' <作成者>米井 優顕</作成者>
    ''' <備考></備考>
    ''' <history version ="2" date ="2009.06.08" name ="伊藤 匡明">クライアントの日付書式に影響されるので日付変換にカルチャを指定</history>
    ''' <history version ="3" date ="2010.07.26" name="HANH-LV">対応内容：現場システムオフライン対応</history>
    ''' <history version ="4" date ="2017.08.24" name="HAI-NM">SQLをPROCES.S5からSERVER2016(プロジェクト武蔵)に移行する。</history>
    <DebuggerStepThrough()>
    Public Shared Function fSqlStr(ByVal pSourceString As Object, ByVal pSourceStringType As BrowsefSqlStr) As String
        Dim xRet As String
        If IsDBNull(pSourceString) Or pSourceString Is Nothing Then
            Return "NULL"
        End If
        Select Case pSourceStringType
            Case BrowsefSqlStr.fSqlStr_Currency, BrowsefSqlStr.fSqlStr_Decimal, BrowsefSqlStr.fSqlStr_Double, BrowsefSqlStr.fSqlStr_Float,
                 BrowsefSqlStr.fSqlStr_Integer, BrowsefSqlStr.fSqlStr_Long, BrowsefSqlStr.fSqlStr_Numeric, BrowsefSqlStr.fSqlStr_Single
                xRet = CStr(pSourceString)
            Case BrowsefSqlStr.fSqlStr_DATE, BrowsefSqlStr.fSqlStr_DateTime, BrowsefSqlStr.fSqlStr_Time
                xRet = String.Format("CONVERT(DATETIME2, '{0}', 126)",
                                CDate(pSourceString).ToString("s", System.Globalization.DateTimeFormatInfo.InvariantInfo))
            Case BrowsefSqlStr.fSqlStr_Text, BrowsefSqlStr.fSqlStr_Memo, BrowsefSqlStr.fSqlStr_LongBinary
                xRet = "N'" & CStr(fReplaceString(fToText(pSourceString), "'", "`")) & "'"
            Case Else
                xRet = "N'" & CStr(fReplaceString(fToText(pSourceString), "'", "`")) & "'"
        End Select
        fSqlStr = xRet
    End Function

    ''' <summary>
    ''' fSqlStr
    ''' </summary>
    ''' <param name="pKeyCode">変換対象変数</param>
    ''' <returns>変換後文字列</returns>
    ''' <remarks></remarks>
    ''' <history version ="1" date ="2010.04.08" name ="米井 優顕">シグネーチャの追加（内部でfsqlstrTypeの呼び出し）</history>
    Public Shared Function fSqlStr(ByVal pKeyCode As Object) As String
        Return fSqlStrType(pKeyCode)
    End Function

    ''' <summary>
    ''' fSqlStrType
    ''' </summary>
    ''' <param name="pKeyCode">編集対象文字</param>
    ''' <returns></returns>
    ''' <remarks>指定型により、fSqlStrを設定後文字列を返却</remarks>
    ''' <history version ="1" date ="2010.03.04" name="米井 優顕">新規作成</history>
    <Browsable(True), Description("明示的に型指定された項目をfSqlStrをかけて返却")>
    Public Shared Function fSqlStrType(ByVal pKeyCode As Object) As String
        Dim xTxt As String
        If TypeOf pKeyCode Is String Then       ''文字列
            xTxt = fSqlStr(pKeyCode, BrowsefSqlStr.fSqlStr_Text)
        ElseIf TypeOf pKeyCode Is Date Then     ''日付
            xTxt = fSqlStr(pKeyCode, BrowsefSqlStr.fSqlStr_DATE)
        Else                                    ''それ以外は数値
            xTxt = fSqlStr(pKeyCode, BrowsefSqlStr.fSqlStr_Numeric)
        End If
        Return xTxt

    End Function

    ''' <summary>fDtSqlStr</summary>
    ''' <param name="pSourceString">元変数</param>
    ''' <param name="pSourceStringType">変数タイプ</param>
    ''' <returns></returns>
    ''' <remarks>DataTable Sql文字列化</remarks>
    ''' <history version ="1" date ="2007.03.05" name ="米井 優顕">V4fJetSqlStrより移行</history>
    ''' <history version ="2" date ="2008.10.01" name ="米井 優顕">SQL文字列の規約がMDBとDataTableで違うために処理を改修</history>
    ''' <history version ="3" date ="2009.03.25" name ="千葉 友則">引数判定にNothing判定を追加</history>
    ''' <history version ="4" date ="2009.06.08" name ="伊藤 匡明">クライアントの日付書式に影響されるので日付変換にカルチャを指定</history>
    ''' <history version ="5" date ="2014.03.12" name ="内田 久義">時刻の書式指定が不正な為、12時間制でフォーマットされる</history>
    <DebuggerStepThrough(), Description("DataTable Sql文字列化")>
    Public Shared Function fDtSqlStr(ByVal pSourceString As Object, ByVal pSourceStringType As BrowsefSqlStr) As String
        Dim xRet As String

        If IsDBNull(pSourceString) Or pSourceString Is Nothing Then
            xRet = "NULL"
        Else
            Select Case pSourceStringType
                Case BrowsefSqlStr.fSqlStr_DateTime
                    ''2009.06.08 UPD START ITO クライアントの日付書式に影響されるのでカルチャを指定
                    'xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy hh:mm:ss") & "#"
                    ''2014.03.12 UPD START
                    'xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy hh:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) & "#"
                    xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) & "#"
                    ''2014.03.12 UPD END
                    ''2009.06.08 UPD END ITO
                Case BrowsefSqlStr.fSqlStr_DATE
                    ''2009.06.08 UPD START ITO クライアントの日付書式に影響されるのでカルチャを指定
                    'xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy") & "#"
                    xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo) & "#"
                    ''2009.06.08 UPD END ITO
                Case BrowsefSqlStr.fSqlStr_Text, BrowsefSqlStr.fSqlStr_Memo, BrowsefSqlStr.fSqlStr_LongBinary
                    Dim dT As New DataTable
                    dT.Columns.Add("STR", GetType(String))
                    dT.Rows.Add()
                    dT.Rows(0).Item("STR") = pSourceString
                    xRet = "'" & CStr(fReplaceString(fToText(dT.Rows(0).Item("STR")), "'", "`")) & "'"
                Case Else
                    xRet = Format$(pSourceString)
            End Select
        End If
        Return xRet
    End Function

    ''' <summary>fChangeTilde -- 文字列の中の"～"のコードを変換する</summary>
    ''' <param name="pString">処理する文字列</param>
    ''' <returns>置換後の文字列</returns>
    ''' <remarks>＆H301C を ＆HFF5E に変換する</remarks>
    ''' <作成者>米井 優顕</作成者>
    ''' <備考>Oracle vs Microsoft 非互換</備考>
    <DebuggerStepThrough()>
    Public Shared Function fChangeTilde(ByVal pString As String) As String
        pString = pString.Replace(ChrW(&H301C), ChrW(&HFF5E))
        Return pString
    End Function

    ''' <summary>fCodeformat -- 数値を指定桁数の右詰め文字でもどす｡</summary>
    ''' <param name="pCode">数値</param>
    ''' <param name="pKeta">指定桁数</param>
    ''' <param name="pFormat">フォーマット文字(Format関数準拠)</param>
    ''' <returns>右詰め文字</returns>
    ''' <remarks>指定数値を指定桁数の文字数で文字列として返す</remarks>
    ''' <作成者>金井　恵史</作成者>
    ''' <備考></備考>
    <DebuggerStepThrough()>
    Public Shared Function fCodeformat(ByVal pCode As Object, Optional ByVal pKeta As Integer = 9, Optional ByVal pFormat As String = "") As String
        Dim nP As Decimal = fToNumberEx(pCode)
        Dim xS As String
        Dim ResultStr As String

        ''書式指定がある場合は書式を設定する
        If pFormat = "" Then
            xS = Format$(nP)
        Else
            xS = Format$(nP, pFormat)
        End If

        ''指定桁数にまとめる
        Select Case pKeta
            Case Len(xS)
                ResultStr = xS
            Case Is > Len(xS)
                ResultStr = Space$(pKeta - Len(xS)) & xS
            Case Else
                ResultStr = Right$(xS, pKeta)
        End Select

        Return ResultStr

    End Function

    ''' <summary>SiuFormat -- 和暦対応フォーマット関数</summary>
    ''' <param name="Expression">変換対象オブジェクト</param>
    ''' <param name="Style">変換文字フォーマット</param>
    ''' <returns>変換後文字列</returns>
    ''' <remarks></remarks>
    ''' <history version ="2" date ="2009.06.08" name ="伊藤 匡明">クライアントの日付書式に影響されるので日付変換にカルチャを指定</history>
    ''' <history version ="3" date ="2009.08.13" name ="内田 久義">年なしの日付書式にもカルチャを指定</history>
    ''' <history version ="4" date ="2009.08.19" name ="千葉 友則">西暦指定時にカレント書式を使用するように変更</history>
    ''' <history version ="5" date ="2009.08.20" name ="千葉 友則">ver4を無効にして、リメイク</history>
    ''' <history version ="6" date ="2010.05.26" name ="清水 勝也">(課題表:3091)SiuFormatの和暦の扱いでエラーが無限ループするのを修正</history>
    <DebuggerStepThrough()>
    Public Shared Function SiuFormat(ByVal Expression As Object, Optional ByVal Style As String = "") As String
        Dim cul As New CultureInfo("ja-JP")
        cul.DateTimeFormat.Calendar = New JapaneseCalendar
        Dim ExpressionDate As Date
        Dim strReturn As String = ""
        Dim strEra As String = ""

        Dim eraKanji() As String = {"", "明", "大", "昭", "平"}
        Dim eraAlpha() As String = {"", "M", "T", "S", "H"}

        If IsDate(Expression) Then
            Expression = CType(Expression, Date)
        End If

        If TypeOf (Expression) Is Date Then ''日付の場合は
            Try
                ExpressionDate = DirectCast(Expression, Date)
                '' 2009.08.20 UPD START
                '    If Style Like "*gggee*" Then                            '平成00'
                '        Style = Style.Replace("gggee", "ggyy")
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '    ElseIf Style Like "*ggge*" Then                         '平成0'
                '        Style = Style.Replace("ggge", "ggy")
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '    ElseIf Style Like "*ggee*" Then                         '平00'
                '        Style = Style.Replace("ggee", "yy")
                '        ''元号取得
                '        strEra = eraKanji(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    ElseIf Style Like "*gge*" Then                          '平0'
                '        Style = Style.Replace("gge", "y")
                '        ''元号取得
                '        strEra = eraKanji(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    ElseIf Style Like "*gee*" Then                          'H00'
                '        Style = Style.Replace("gee", "yy")
                '        ''元号取得
                '        strEra = eraAlpha(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    ElseIf Style Like "*ge*" Then                           'H0'
                '        Style = Style.Replace("ge", "y")
                '        ''元号取得
                '        strEra = eraAlpha(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    Else
                '        ''2009.06.08 UPD START ITO クライアントの日付書式に影響されるのでカルチャを指定
                '        'strReturn = Format(Expression, Style)
                '        ''2009.08.13 UPD START
                '        'strReturn = ExpressionDate.ToString(Style, System.Globalization.DateTimeFormatInfo.InvariantInfo)
                '        ''2009.08.19 UPD START
                '        'strReturn = ExpressionDate.ToString(Style, cul)
                '        'If Style.IndexOf("yyyy") > -1 Then
                '        '    'strReturn = ExpressionDate.ToString("yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo) & ExpressionDate.ToString(Style.Replace("yyyy", ""), cul)
                '        '    strReturn = ExpressionDate.ToString(Style, System.Globalization.DateTimeFormatInfo.CurrentInfo)
                '        'Else
                '        '    strReturn = ExpressionDate.ToString(Style, cul)
                '        'End If
                '        ''2009.08.19 UPD END
                '        ''2009.08.13 UPD END
                '        ''2009.06.08 UPD END ITO
                '    End If

                ''2010.05.26 ADD START SHIMIZU
                ' 曜日部分はカルチャに関わらず和暦とするので先に置換
                Style = ReplaceJapaneseDayOfWeek(ExpressionDate, Style)
                ''2010.05.26 ADD END

                '' まず「g」、「e」が含まれるか否かでカルチャ指定するか否かを判断する
                If Style.ToLower.IndexOf("g") > -1 OrElse Style.ToLower.IndexOf("e") > -1 Then
                    '' カルチャ指定

                    ''「e」⇒「y」
                    Style = Style.Replace("e"c, "y"c)

                    '' 元号書式にて、「g」の数を取得
                    Dim eraCount As Integer = 0
                    For Each c As Char In Style.ToCharArray
                        If c.Equals("g"c) Then
                            eraCount += 1
                        End If
                    Next

                    '' 標準で対応できない書式は、配列から文字取得
                    Select Case eraCount
                        Case 1
                            ''「g」消去
                            Style = Style.Replace("g"c, "")
                            ''元号アルファベット取得
                            strEra = eraAlpha(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                        Case 2
                            ''「g」消去
                            Style = Style.Replace("g"c, "")
                            ''元号漢字略取得
                            strEra = eraKanji(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                        Case Else
                            strEra = String.Empty
                    End Select

                    '' 元号 + 標準で戻す（元号文字列がない場合は標準だけで対応）
                    strReturn = strEra & ExpressionDate.ToString(Style, cul)

                Else
                    ''2010.05.26 DEL START SHIMIZU 和暦での例外時に無限ループの原因になっていたので削除
                    ' '' ※曜日はかならず日本書式で返す
                    'Dim ddd As String = String.Empty
                    ' '' 曜日指定書式文字を、あらかじめカルチャ指定で取得し書き換えておく
                    'If Style.ToLower.IndexOf("dddd") > -1 Then
                    '    ddd = ExpressionDate.ToString("dddd", cul)
                    '    Style = Style.Replace("dddd", ddd)
                    'ElseIf Style.ToLower.IndexOf("ddd") > -1 Then
                    '    ddd = ExpressionDate.ToString("ddd", cul)
                    '    Style = Style.Replace("ddd", ddd)
                    'End If
                    ''2010.05.26 DEL END

                    '' カルチャ無指定
                    strReturn = ExpressionDate.ToString(Style, System.Globalization.DateTimeFormatInfo.InvariantInfo)
                End If
                '' 2009.08.20 UPD END

            Catch ex As Exception   ''失敗したらそのまま
                '' 2009.05.19 ADD START
                ' 失敗時は元号取得できないと判断し、西暦で戻す
                ' 年書式を一旦削除し、西暦4桁指定
                Style = Style.Replace("g"c, "")
                Style = Style.Replace("e"c, "")
                Style = Style.Replace("y"c, "")
                Style = "yyyy" & Style
                '' 2009.05.19 ADD END
                ''2009.06.08 UPD START ITO クライアントの日付書式に影響されるのでカルチャを指定
                'strReturn = Format(Expression, Style)
                ''2009.08.13 UPD START
                'strReturn = ExpressionDate.ToString(Style, System.Globalization.DateTimeFormatInfo.InvariantInfo)
                ''2009.08.19 UPD START
                'strReturn = ExpressionDate.ToString(Style, cul)
                ''2009.08.20 DEL START
                'If Style.IndexOf("yyyy") > -1 Then
                '    'strReturn = ExpressionDate.ToString("yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo) & ExpressionDate.ToString(Style.Replace("yyyy", ""), cul)
                '    strReturn = ExpressionDate.ToString(Style, System.Globalization.DateTimeFormatInfo.CurrentInfo)
                'Else
                '    strReturn = ExpressionDate.ToString(Style, cul)
                'End If
                ''2009.08.20 DEL END
                ''2009.08.19 UPD END
                ''2009.08.13 UPD END
                ''2009.06.08 UPD END ITO

                ''2009.08.20 ADD 
                '' 書式を書き換えて、再実行
                strReturn = SiuFormat(Expression, Style)
            End Try
        Else
            strReturn = Format(Expression, Style)
        End If

        Return strReturn
    End Function

    ''' <summary>
    ''' ReplaceJapaneseDayOfWeek -- 曜日書式を和暦で置換
    ''' </summary>
    ''' <param name="targetDate">対象日付</param>
    ''' <param name="style">書式</param>
    ''' <returns>置換した書式</returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.05.26" name="清水 勝也">新規作成</history>
    <DebuggerStepThrough()>
    Public Shared Function ReplaceJapaneseDayOfWeek(ByVal targetDate As Date, ByVal style As String) As String
        Dim xDay As String  ' 曜日変数

        ' 曜日を取得して変数に日本語を設定
        Select Case targetDate.DayOfWeek
            Case DayOfWeek.Sunday
                xDay = "日"
            Case DayOfWeek.Monday
                xDay = "月"
            Case DayOfWeek.Tuesday
                xDay = "火"
            Case DayOfWeek.Wednesday
                xDay = "水"
            Case DayOfWeek.Thursday
                xDay = "木"
            Case DayOfWeek.Friday
                xDay = "金"
            Case DayOfWeek.Saturday
                xDay = "土"
            Case Else
                xDay = String.Empty
        End Select

        ' 書式に合わせて置換
        If style.IndexOf("dddd") > -1 Then
            style = style.Replace("dddd", xDay & "曜日")
        ElseIf style.IndexOf("ddd") > -1 Then
            style = style.Replace("ddd", xDay)
        End If

        Return style
    End Function

    ''' <summary>ConvCrLf -- vbCrLf←→\n</summary>
    ''' <param name="str">変換対象文字列</param>
    ''' <param name="mode">0:vbCrLf→\n 1:\n→vbCrLf</param>
    ''' <returns>変換後文字列</returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2008.05.12" name="千葉 友則">新規作成</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    Public Shared Function ConvCrLf(ByVal str As String, ByVal mode As Integer) As String

        Dim xPattern As String
        Dim xReplace As String

        If str Is Nothing Then
            Return String.Empty
        End If

        If mode = 0 Then
            xPattern = vbCrLf
            xReplace = "\n"
        Else
            xPattern = "\\n"
            xReplace = vbCrLf
        End If

        '' 正規表現で検索
        Dim regex As New System.Text.RegularExpressions.Regex(xPattern)

        '' パターンマッチ
        If Not regex.IsMatch(str) Then
            Return str
        End If

        '' 置換
        Return regex.Replace(str, xReplace)

    End Function

    ''' <summary>
    ''' NumDateFormat
    ''' </summary>
    ''' <param name="strTextChk"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    Public Shared Function NumDateFormat(ByVal strTextChk As String) As String
        If strTextChk Is Nothing Then
            strTextChk = String.Empty
        End If
        If Len(strTextChk) = 6 Then
            '西暦下２桁＋月(2桁)＋日(2桁)と判断する。
            strTextChk = CDate(String.Format("{0:0000/00/00}", CDbl("20" & strTextChk)))
        ElseIf Len(strTextChk) = 8 Then
            '西暦４桁＋月(2桁)＋日(2桁)と判断する。
            strTextChk = CDate(String.Format("{0:0000/00/00}", CDbl(strTextChk)))
        End If

        Return strTextChk

    End Function

    ''' <summary>
    ''' fSqlStr -- DataColumnから型やPROCESSでの用途を判断してOracle形式のSQL値に変換
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="col"></param>
    ''' <param name="updSyubetu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="清水 勝也">新規作成</history>
    ''' <history version="2" date="2017.08.24" name="HAI-NM">SQLをPROCES.S5からSERVER2016(プロジェクト武蔵)に移行する。</history>
    Public Shared Function fSqlStr(ByVal value As Object, ByVal col As DataColumn, ByVal updSyubetu As UpdType) As String
        Dim xSqlValue As String

        Select Case col.ColumnName
            Case "UD_DATE"
                xSqlValue = "SYSDATETIME()"
            Case "AD_DATE"
                If updSyubetu = UpdType.Insert Then
                    xSqlValue = "SYSDATETIME()"
                Else
                    If fToDate(value) = SysMinDate Then
                        xSqlValue = "SYSDATETIME()"
                    Else
                        xSqlValue = fSqlStr(value, BrowsefSqlStr.fSqlStr_DateTime)
                    End If
                End If
            Case "OPID"
                xSqlValue = fSqlStr(SysOpid, BrowsefSqlStr.fSqlStr_Text)
            Case "UD_USER"
                xSqlValue = fSqlStr(SysUserNo, BrowsefSqlStr.fSqlStr_Integer)
            Case "RIREKI_NO", "RIREKI_INDEX"
                xSqlValue = fSqlStr(fToDec(value), BrowsefSqlStr.fSqlStr_Decimal)
            Case "YOBI_DATE1", "YOBI_DATE2", "YOBI_DATE3", "YOBI_DATE4", "YOBI_DATE5",
                 "YOBI_DATE6", "YOBI_DATE7", "YOBI_DATE8", "YOBI_DATE9", "YOBI_DATE10"
                If fToDate(value) = SysMinDate Then
                    xSqlValue = "NULL"
                Else
                    xSqlValue = fSqlStr(value, BrowsefSqlStr.fSqlStr_DateTime)
                End If
            Case Else
                If IsDBNull(value) Then
                    xSqlValue = "NULL"
                Else
                    ' 型判断
                    xSqlValue = fSqlStrType(value)
                End If
        End Select

        Return xSqlValue
    End Function

#End Region

#Region " SQL生成関数 "

    ''' <summary>
    ''' fSqlStrInsert -- DataRowからInsert文を生成
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="清水 勝也">新規作成</history>
    Public Shared Function fSqlStrInsert(ByVal tableName As String, ByVal source As DataRow) As String
        ' 更新種別INSERTでSQL生成
        Return fSqlStrInsert(tableName, source, UpdType.Insert)
    End Function

    ''' <summary>
    ''' fSqlStrInsert -- DataRowからInsert文を生成
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <param name="updSyubetu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="清水 勝也">新規作成</history>
    Public Shared Function fSqlStrInsert(ByVal tableName As String, ByVal source As DataRow, ByVal updSyubetu As UpdType) As String
        Dim sbSqlTxt As New System.Text.StringBuilder    ' SQL
        Dim sbFieldsTxt As New System.Text.StringBuilder ' フィールド
        Dim sbValuesTxt As New System.Text.StringBuilder ' 値

        sbSqlTxt.AppendLine("INSERT ")
        sbSqlTxt.AppendLine("INTO ")
        sbSqlTxt.AppendLine("		{0} ")
        sbSqlTxt.AppendLine("( ")
        sbSqlTxt.Append("{1} ")
        sbSqlTxt.AppendLine(") ")
        sbSqlTxt.AppendLine("VALUES ")
        sbSqlTxt.AppendLine("( ")
        sbSqlTxt.Append("{2} ")
        sbSqlTxt.AppendLine(") ")

        For Each col As DataColumn In source.Table.Columns
            sbFieldsTxt.AppendFormat(",		{0} ", col.ColumnName).AppendLine()
            sbValuesTxt.AppendFormat(",		{0} ", fSqlStr(source(col.ColumnName), col, updSyubetu)).AppendLine()
        Next

        Return String.Format(sbSqlTxt.ToString,
                             tableName,
                             sbFieldsTxt.Remove(0, 1).ToString,
                             sbValuesTxt.Remove(0, 1).ToString
                             )
    End Function

    ''' <summary>
    ''' fSqlStrUpdate -- DataRowとKey配列からUpdate文を生成
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <param name="keyColNames"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="清水 勝也">新規作成</history>
    Public Shared Function fSqlStrUpdate(ByVal tableName As String, ByVal source As DataRow, ByVal keyColNames() As String) As String
        Dim sbSqlTxt As New System.Text.StringBuilder    ' SQL
        Dim sbValuesTxt As New System.Text.StringBuilder ' 値
        Dim sbWhereTxt As New System.Text.StringBuilder  ' 条件

        sbSqlTxt.AppendLine("UPDATE ")
        sbSqlTxt.AppendLine("		{0} ")
        sbSqlTxt.AppendLine("SET ")
        sbSqlTxt.Append("{1} ")
        sbSqlTxt.AppendLine("WHERE ")
        sbSqlTxt.Append("{2} ")

        For Each col As DataColumn In source.Table.Columns
            sbValuesTxt.AppendFormat(",		{0}	=	{1} ",
                                     col.ColumnName,
                                     fSqlStr(source.Item(col.ColumnName), col, UpdType.Update)
                                     ).AppendLine()
        Next

        For Each keyColName As String In keyColNames
            sbWhereTxt.AppendFormat("AND		{0}	=	{1} ",
                                    keyColName,
                                    fSqlStr(source.Item(keyColName), source.Table.Columns(keyColName), UpdType.Update)
                                    ).AppendLine()
        Next

        Return String.Format(sbSqlTxt.ToString,
                             tableName,
                             sbValuesTxt.Remove(0, 1).ToString,
                             sbWhereTxt.Remove(0, 3).ToString
                             )
    End Function

    ''' <summary>
    ''' fSqlStrDelete -- DataRowとKey配列からDelete文を生成
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <param name="keyColNames"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="清水 勝也">新規作成</history>
    Public Shared Function fSqlStrDelete(ByVal tableName As String, ByVal source As DataRow, ByVal keyColNames() As String) As String
        Dim sbSqlTxt As New System.Text.StringBuilder    ' SQL
        Dim sbWhereTxt As New System.Text.StringBuilder  ' 条件

        sbSqlTxt.AppendLine("DELETE ")
        sbSqlTxt.AppendLine("FROM ")
        sbSqlTxt.AppendLine("		{0} ")
        sbSqlTxt.AppendLine("WHERE ")
        sbSqlTxt.Append("{1} ")

        For Each keyColName As String In keyColNames
            sbWhereTxt.AppendFormat("AND		{0}	=	{1} ",
                                    keyColName,
                                    fSqlStr(source.Item(keyColName), source.Table.Columns(keyColName), UpdType.Delete)
                                    ).AppendLine()
        Next

        Return String.Format(sbSqlTxt.ToString,
                             tableName,
                             sbWhereTxt.Remove(0, 3).ToString
                             )
    End Function

#End Region

#Region " 抽出・比較関数 "
    ''***************************************************************************************************************************************************

    ''' <summary>fLeftByte -- 左部文字列抽出</summary>
    ''' <param name="pSource">対象文字列</param>
    ''' <param name="pLeftLen">抽出バイト数</param>
    ''' <param name="pCharSet">エンコード指定: 1=shift-jis 2=Unicode(16) 3=UTF8 4=UTF7 5=UTF32</param>
    ''' <returns>抽出文字列</returns>
    ''' <remarks>指定文字列の左から指定バイト数を抽出する</remarks>
    ''' <作成者>Process4からの複製作成</作成者>
    ''' <備考>全角文字の半分で終了する場合は切り捨てる</備考>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    <DebuggerStepThrough()>
    Public Shared Function fLeftByte(ByVal pSource As Object, ByVal pLeftLen As Integer, ByVal pCharSet As Integer) As String
        Dim xBuf As String = ""
        Dim xBuf2 As String = ""
        Dim i As Integer

        Try
            ''Nullの場合はNullで返信
            If pSource Is DBNull.Value Then
                Return DBNull.Value.ToString
            End If

            If pSource Is Nothing Then
                Return String.Empty
            End If

            ''長さゼロの場合は空文字返信
            If CType(pSource, String).Length = 0 Then
                Return ""
            End If

            ''文字列チェック
            For i = 1 To Len(pSource) Step 1
                xBuf2 = xBuf & Mid$(CType(pSource, String), i, 1)

                'SJISでのバイト数との比較
                If LenB(xBuf2, pCharSet) > pLeftLen Then
                    Exit For
                Else
                    xBuf = xBuf2
                End If
            Next i
            Return xBuf

        Catch ex As Exception

            Return ""

        End Try

    End Function

    ''' <summary>fStrCompSJIS -- 文字列をSJISで比較する</summary>
    ''' <param name="pxFrom">開始文字列</param>
    ''' <param name="pxTo">終了文字列</param>
    ''' <returns>
    ''' True  :開始文字列 ＜= 終了文字列
    ''' False :開始文字列 ＞ 終了文字列
    ''' </returns>
    ''' <remarks>文字列をSJISで比較する</remarks>
    ''' <作成者>Process4からの複製作成</作成者>
    ''' <備考></備考>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    <DebuggerStepThrough()>
    Public Shared Function fStrCompSJIS(ByVal pxFrom As String, ByVal pxTo As String) As Boolean

        Dim nFrom() As Byte
        Dim nTo() As Byte
        Dim nStart As Integer
        Dim nEnd As Integer
        Dim bEQ As Boolean
        Dim i As Integer

        If pxFrom Is Nothing Then
            pxFrom = String.Empty
        End If
        If pxTo Is Nothing Then
            pxTo = String.Empty
        End If

        'S-JISへのエンコードオブジェクト
        Dim SJIS_Encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        '' 変数の初期化
        fStrCompSJIS = False
        nFrom = SJIS_Encode.GetBytes(pxFrom)
        'nFrom = StrConv(pxFrom, vbFromUnicode)
        nTo = SJIS_Encode.GetBytes(pxTo)
        'nTo = StrConv(pxTo, vbFromUnicode)

        '' 比較の開始ポインタ
        If LBound(nFrom) < LBound(nTo) Then
            nStart = LBound(nTo)
        Else
            nStart = LBound(nFrom)
        End If

        '' 比較の終了ポインタ・全て一致の結果値
        If UBound(nFrom) > UBound(nTo) Then
            nEnd = UBound(nTo)
            bEQ = False
        Else
            nEnd = UBound(nFrom)
            bEQ = True
        End If

        '' 比較ループ
        For i = nStart To nEnd
            If nFrom(i) > nTo(i) Then
                fStrCompSJIS = False
                Exit Function
            ElseIf nFrom(i) < nTo(i) Then
                fStrCompSJIS = True
                Exit Function
            End If
        Next

        '' 全て一致の結果値
        fStrCompSJIS = bEQ

    End Function

    ''' <summary>SplitString -- 文字列を一時ずつ分解し配列に落としこむ</summary>
    ''' <param name="str">対象文字列</param>
    ''' <returns></returns>
    ''' <remarks>Splitのセパレーターが無い版</remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    <DebuggerStepThrough()>
    Public Shared Function SplitString(ByVal str As String) As String()
        'Dim strLength As Integer
        'Dim strArray(0) As String

        'strLength = str.Length
        'For i As Integer = 0 To strLength - 1 Step 1
        '    ReDim Preserve strArray(i)
        '    strArray(i) = Mid(str, i + 1, 1)
        'Next i

        'Return strArray

        If str Is Nothing Then
            str = String.Empty
        End If

        Return Array.ConvertAll(str.ToCharArray, New Converter(Of Char, String)(AddressOf Char.ToString))

    End Function

    ''' <summary>LenB -- バイト数取得</summary>
    ''' <param name="Expression">カウント対象文字列</param>
    ''' <returns>カウント結果</returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Shared Function LenB(ByVal Expression As String) As Integer
        Return LenB(Expression, 1)
    End Function

    ''' <summary>LenB -- バイト数取得</summary>
    ''' <param name="Expression">カウント対象文字列</param>
    ''' <param name="pCodingNumber">エンコード指定: 1=shift-jis 2=Unicode(16) 3=UTF8 4=UTF7 5=UTF32</param>
    ''' <returns>カウント結果</returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    <DebuggerStepThrough()>
    Public Shared Function LenB(ByVal Expression As String, ByVal pCodingNumber As Integer) As Integer
        If Expression Is Nothing Then
            Expression = String.Empty
        End If

        Dim sEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift-jis")
        Select Case pCodingNumber
            Case 0

            Case 1      ''Shift-Jis
                Return sEncoding.GetByteCount(Expression)
            Case 2      ''Unicode(16) 
                Return Encoding.Unicode.GetByteCount(Expression)
            Case 3      ''UTF8
                Return Encoding.UTF8.GetByteCount(Expression)
            Case 4      ''UTF7
                Return Encoding.UTF7.GetByteCount(Expression)
            Case 5      ''UTF32
                Return Encoding.UTF32.GetByteCount(Expression)
        End Select

    End Function

    ''' <summary>SysCodeEquals -- 各種システムコードの最小/最大比較</summary>
    ''' <param name="value">比較対象値</param>
    ''' <returns>同値の場合はTrue</returns>
    ''' <remarks>SysMinCodeとの比較</remarks>
    ''' <history version="1" date="2008.05.23" name="千葉 友則">新規作成</history>
    Public Shared Function SysCodeEquals(ByVal value As Object) As Boolean
        Return SysCodeEquals(value, ControlKbn.Only, SysCodeType.Code)
    End Function

    ''' <summary>SysCodeEquals -- 各種システムコードの最小/最大比較</summary>
    ''' <param name="value">比較対象値</param>
    ''' <param name="conkbn">ControlKbn列挙体</param>
    ''' <returns>同値の場合はTrue</returns>
    ''' <remarks>SysMinCode/SysMaxCodeとの比較</remarks>
    ''' <history version="1" date="2008.05.23" name="千葉 友則">新規作成</history>
    Public Shared Function SysCodeEquals(ByVal value As Object, ByVal conkbn As Integer) As Boolean
        Return SysCodeEquals(value, conkbn, SysCodeType.Code)
    End Function

    ''' <summary>SysCodeEquals -- 各種システムコードの最小/最大比較</summary>
    ''' <param name="value">比較対象値</param>
    ''' <param name="conkbn">ControlKbn列挙体</param>
    ''' <param name="type">SysCodeType列挙体</param>
    ''' <returns>同値の場合はTrue</returns>
    ''' <remarks>typeで指定されたSysMin***/SysMax***との比較</remarks>
    ''' <history version="1" date="2008.05.23" name="千葉 友則">新規作成</history>
    ''' <history version="2" date="2009.03.02" name="大内 登志也">部門コードの文字コード化対応</history>
    Public Shared Function SysCodeEquals(ByVal value As Object, ByVal conkbn As Integer, ByVal type As SysCodeType) As Boolean

        Dim MinValue As Object
        Dim MaxValue As Object
        Dim CompValue As Object

        '' タイプ別に最小/最大を取得
        Select Case type
            Case SysCodeType.Code           '' コード
                MinValue = SysMinCode
                MaxValue = SysMaxCode
            Case SysCodeType.Date           '' 日付
                MinValue = SysMinDate
                MaxValue = SysMaxDate
            Case SysCodeType.SKoj           '' 集計工事コード
                MinValue = SysMinSKoj
                MaxValue = SysMaxSKoj
            Case SysCodeType.Koj            '' 工事コード
                MinValue = SysMinKoj
                MaxValue = SysMaxKoj
            Case SysCodeType.KojUw          '' 工事詳細コード
                MinValue = SysMinKojUw
                MaxValue = SysMaxKojUw
            Case SysCodeType.Tor            '' 取引先コード
                MinValue = SysMinTor
                MaxValue = SysMaxTor
            Case SysCodeType.Soko           '' 倉庫コード
                MinValue = SysMinSoko
                MaxValue = SysMaxSoko
            Case SysCodeType.Syohin         '' 商品コード
                MinValue = SysMinSyohin
                MaxValue = SysMaxSyohin
            Case SysCodeType.Jyu            '' 受注コード
                MinValue = SysMinJyu
                MaxValue = SysMaxJyu
            Case SysCodeType.Hac            '' 発注コード
                MinValue = SysMinHac
                MaxValue = SysMaxHac
            Case SysCodeType.Syu            '' 出荷コード
                MinValue = SysMinSyu
                MaxValue = SysMaxSyu
            Case SysCodeType.Uri            '' 売上コード
                MinValue = SysMinUri
                MaxValue = SysMaxUri
            Case SysCodeType.Nyu            '' 入荷コード
                MinValue = SysMinNyu
                MaxValue = SysMaxNyu
            Case SysCodeType.Sir            '' 仕入コード
                MinValue = SysMinSir
                MaxValue = SysMaxSir
                ''2009.03.02 ADD START OOUCHI 部門コードの文字コード化対応
            Case SysCodeType.Bmn
                MinValue = SysMinBmn
                MaxValue = SysMaxBmn
                ''2009.03.02 ADD END
            Case Else                       '' その他はコードで処理
                MinValue = SysMinCode
                MaxValue = SysMaxCode
        End Select

        '' 比較する区分により、比較値を取得
        If (conkbn = ControlKbn.Only) Or (conkbn = ControlKbn.From) Then
            '' 単独または範囲開始との確認はMinValueで確認
            CompValue = MinValue
        Else
            '' 以外はMaxValueで確認
            CompValue = MaxValue
        End If

        '' 比較
        Return value.Equals(CompValue)

    End Function

#End Region

#Region " 算出関数 "
    ''***************************************************************************************************************************************************

    ''' <summary>
    ''' fStartDate--月度開始日算出
    ''' </summary>
    ''' <param name="pDdate">月度</param>
    ''' <param name="pSIME">締日</param>
    ''' <returns>月度開始日</returns>
    ''' <remarks>指定月度の開始日を月度＋締日で算出する。</remarks>
    ''' <作成者>Process4からの複製作成</作成者>
    ''' <備考></備考>
    <DebuggerStepThrough()>
    Public Shared Function fStartDate(ByVal pDdate As Object, ByVal pSIME As Integer) As Date
        Dim dGaitouDate As Object
        Dim dGaitouWk As Object

        ''初期値Null
        fStartDate = SysMinDate

        ''日付以外だった場合は終了
        If IsDate(pDdate) = False Then
            fStartDate = SysMinDate
            Exit Function
        End If

        ''該当日の日を１に強制固定(月度にする)
        dGaitouDate = DateSerial(Year(pDdate), Month(pDdate), 1)

        ''閉め日が３１の場合は１日が該当なのでそのまま戻す｡
        If pSIME = 31 Then
            fStartDate = dGaitouDate
            Exit Function
        End If

        ''閉め日が３１の場合は１日が該当なのでそのまま戻す｡
        If Microsoft.VisualBasic.Day(DateSerial(Year(dGaitouDate), Month(dGaitouDate) + 1, 1 - 1)) < pSIME Then
            If Microsoft.VisualBasic.Day(DateSerial(Year(dGaitouDate), Month(dGaitouDate), 1 - 1)) > pSIME Then
                fStartDate = DateSerial(Year(dGaitouDate), Month(dGaitouDate) - 1, pSIME + 1)
                Exit Function
            End If
        End If
        dGaitouWk = fEndDate(dGaitouDate, pSIME)
        dGaitouWk = DateAdd("m", -1, dGaitouWk)
        fStartDate = DateAdd("d", 1, dGaitouWk)
    End Function

    ''' <summary>
    ''' fEndDate--月度末日算出
    ''' </summary>
    ''' <param name="pDdate">月度</param>
    ''' <param name="pSIME">締日</param>
    ''' <returns>月度末日</returns>
    ''' <remarks>指定月度の最終日を月度＋締日で算出する。</remarks>
    ''' <作成者>Process4からの複製作成</作成者>
    ''' <備考></備考>
    <DebuggerStepThrough()>
    Public Shared Function fEndDate(ByVal pDdate As Object, ByVal pSIME As Integer) As Date
        Dim dGaitouDate As Object
        Dim dGaitoudd As Integer

        ''初期値=SysMinDate
        fEndDate = SysMinDate

        ''日付以外だった場合は終了
        If IsDate(pDdate) = False Then
            Exit Function
        End If

        ''該当日の日を１に強制固定(月度にする)
        dGaitouDate = DateSerial(Year(pDdate), Month(pDdate), 1)

        ''締日より最終日を算出
        If IsNumeric(pSIME) = False Then
            Exit Function
        ElseIf pSIME = 31 Then
            dGaitoudd = 31
        ElseIf Microsoft.VisualBasic.Day(DateSerial(Year(pDdate), Month(pDdate) + 1, 1 - 1)) < pSIME Then
            dGaitoudd = 31
        ElseIf (pSIME >= 1) And (pSIME < 31) Then
            dGaitoudd = pSIME
        Else
            fEndDate = SysMinDate
            Exit Function
        End If

        ''最終日が31の場合は月別最終日を算出，以外はそのまま代入
        If dGaitoudd = 31 Then
            fEndDate = DateSerial(Year(dGaitouDate), Month(dGaitouDate) + 1, 1 - 1)
        Else
            fEndDate = DateSerial(Year(dGaitouDate), Month(dGaitouDate), dGaitoudd)
        End If
    End Function

    ''' <summary>
    ''' fProcessMonth--月度を算出する
    ''' </summary>
    ''' <param name="pDdate">該当年月日</param>
    ''' <param name="pSIME">締日</param>
    ''' <returns>対象月度</returns>
    ''' <remarks>該当日付がどの月度に属するか戻す</remarks>
    ''' <作成者>金井　恵史</作成者>
    ''' <備考></備考>
    <DebuggerStepThrough()>
    Public Shared Function fProcessMonth(ByVal pDdate As Object, ByVal pSIME As Integer) As Date
        Dim dGaitouDate As Object
        Dim nGaitoudd As Integer
        Dim nGaitouSime As Integer

        ''初期値Null
        fProcessMonth = SysMinDate

        ''日付以外だった場合は終了
        If IsDate(pDdate) = False Then
            fProcessMonth = SysMinDate
            Exit Function
        End If

        ''月度にする
        dGaitouDate = DateSerial(Year(pDdate), Month(pDdate), 1)
        nGaitoudd = Microsoft.VisualBasic.Day(pDdate)

        ''締日より判断
        If pSIME = 31 Then
            nGaitouSime = 31
        ElseIf (pSIME >= 1) And (pSIME < 31) Then
            nGaitouSime = pSIME
        Else
            fProcessMonth = SysMinDate
            Exit Function
        End If

        ''今月か次月か判断
        If nGaitoudd <= nGaitouSime Then
            fProcessMonth = dGaitouDate
        Else
            fProcessMonth = DateSerial(Year(dGaitouDate), Month(dGaitouDate) + 1, 1)
        End If
    End Function

    ''' <summary>fKishuMonth -- 期首年月を出す</summary>
    ''' <param name="pNengetu">画面で指定された年月</param>
    ''' <param name="pStart_YM">事業開始年月</param>
    ''' <returns>期首年月</returns>
    ''' <remarks>画面の年月と事業開始年月から期首の月を求める</remarks>
    ''' <history version="" date="" name="千葉 友則">Process4からの複製</history>
    <DebuggerStepThrough()>
    Public Shared Function fKishuMonth(ByVal pNengetu As Object, ByVal pStart_YM As Object) As Date

        Dim dGaitouDate As Date
        Dim dKariDate As Date

        If IsDate(pNengetu) = False Then
            Return SysMinDate
        End If
        If IsDate(pStart_YM) = False Then
            Return SysMinDate
        End If

        dGaitouDate = DateSerial(Year(pNengetu), Month(pNengetu), 1)
        dKariDate = DateSerial(Year(pNengetu), Month(pStart_YM), 1)

        If dGaitouDate < dKariDate Then
            Return DateAdd("yyyy", -1, dKariDate)
        Else
            Return dKariDate
        End If

    End Function

    ''' <summary>fRound -- 端数処理</summary>
    ''' <param name="pKin">対象金額</param>
    ''' <param name="pHasuu">端数処理区分</param>
    ''' <param name="pTani">丸め単位(1で小数点以下，10で10円以下)</param>
    ''' <returns>計算結果</returns>
    ''' <remarks>指定単位で数値の丸め処理を行います。</remarks>
    ''' <作成者>Process4からの複製作成</作成者>
    ''' <備考></備考>
    <DebuggerStepThrough()>
    Public Shared Function fRound(ByVal pKin As Object, ByVal pHasuu As BrowseHasu, ByVal pTani As Object) As Decimal
        Dim nKin As Decimal
        Dim nTani As Decimal
        Dim nChukan1 As Decimal
        Dim nChukan2 As Decimal

        If IsNumeric(pKin) = False Then
            Return 0
        Else
            nKin = CDec(pKin)
        End If
        If pTani Is Nothing Then
            nTani = 1
        Else
            nTani = CDec(pTani)
        End If

        nChukan1 = nKin / nTani
        Select Case pHasuu
            Case BrowseHasu.bHasu_Kiriage       '切り上げ
                nChukan2 = -System.Math.Sign(nChukan1) * Int(-System.Math.Abs(nChukan1))
            Case BrowseHasu.bHasu_Kirisute      '切り捨て
                nChukan2 = Fix(nChukan1)
            Case BrowseHasu.bHasu_Shisyagonyu   '四捨五入
                nChukan2 = System.Math.Sign(nChukan1) * Int(System.Math.Abs(nChukan1) + 0.5D)
            Case Else
                nChukan2 = 0
        End Select

        Return nChukan2 * nTani

    End Function

    ''' <summary>
    ''' ZeiCalculate
    ''' </summary>
    ''' <param name="pKin">対象金額</param>
    ''' <param name="pZeiInput">税入力区分</param>
    ''' <param name="pZeiRitsu">税率</param>
    ''' <param name="pHasu">端数処理区分</param>
    ''' <returns>税額</returns>
    ''' <remarks></remarks>
    ''' <history version ="1" date ="2010.04.08" name="米井 優顕">新規作成</history>
    <DebuggerStepThrough(), Browsable(True), Description("税算出 pKin:対象金額 pZeiInput:1=外税2=内税 pZeiRitsu:税率 pHasu:端数処理区分")>
    Public Shared Function ZeiCalculate(ByVal pKin As Decimal, ByVal pZeiInput As Integer, ByVal pZeiRitsu As Decimal, ByVal pHasu As Integer) As Decimal

        Dim wZei As Decimal = 0

        Select Case pZeiInput
            Case 1
                wZei = fRound(pKin * pZeiRitsu / 100, pHasu, 1)
            Case 2
                wZei = fRound(pKin * pZeiRitsu / (100 + pZeiRitsu), pHasu, 1)
            Case Else
                wZei = 0
        End Select


        Return wZei

    End Function

#End Region

#Region " 変換関数 "
    ''***************************************************************************************************************************************************

    ''' <summary>fNvl -- Null置き換え</summary>
    ''' <param name="pValue">対象変数</param>
    ''' <param name="pRepValue">Null時置き換え変数</param>
    ''' <returns>Null置き換え後の値</returns>
    ''' <remarks>pValueがNothingまたはNullの場合はpRepValueを，それ以外はpValueをそのまま返す</remarks>
    <DebuggerStepThrough()>
    Public Shared Function fNvl(ByVal pValue As Object, ByVal pRepValue As Object) As Object
        If pValue Is Nothing OrElse pValue Is DBNull.Value Then
            Return pRepValue
        Else
            Return pValue
        End If
    End Function

    ''' <summary>fNvlEx -- Null置き換え</summary>
    ''' <param name="pValue">対象変数</param>
    ''' <param name="pRepValue">Null時置き換え変数</param>
    ''' <remarks>
    ''' pValueがNothingまたはNullの場合はpRepValueを，それ以外はpValueをそのまま返す
    ''' Option Strict = True の対応
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Sub fNvlEx(ByRef pValue As Object, ByVal pRepValue As Object)
        If pValue Is Nothing OrElse pValue Is DBNull.Value Then
            pValue = pRepValue
        End If
    End Sub

    ''' <summary>fToNumber -- 数値化</summary>
    ''' <param name="pVal">対象変数</param>
    ''' <param name="pRetType">返却値の数値ﾀｲﾌﾟ 省略時はDecimalにて返却</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Shared Function fToNumber(ByVal pVal As Object, Optional ByVal pRetType As BrowsefToNumber = BrowsefToNumber.fToNumber_DECIMAL) As Object
        On Error GoTo fToNumberErr
        Dim xRet As Integer
        Dim xErrm As String
        If IsNumeric(pVal) = False Then
            xRet = 0
            Select Case pRetType
                Case BrowsefToNumber.fToNumber_INTEGER
                    fToNumber = CInt(xRet)
                Case BrowsefToNumber.fToNumber_LONG
                    fToNumber = CLng(xRet)
                Case BrowsefToNumber.fToNumber_SINGLE
                    fToNumber = CSng(xRet)
                Case BrowsefToNumber.fToNumber_DOUBLE
                    fToNumber = CDbl(xRet)
                Case BrowsefToNumber.fToNumber_SHORT
                    fToNumber = CShort(xRet)
                Case Else
                    fToNumber = CDec(xRet)
            End Select
        Else
            Select Case pRetType
                Case BrowsefToNumber.fToNumber_INTEGER
                    fToNumber = CInt(pVal)
                Case BrowsefToNumber.fToNumber_LONG
                    fToNumber = CLng(pVal)
                Case BrowsefToNumber.fToNumber_SINGLE
                    fToNumber = CSng(pVal)
                Case BrowsefToNumber.fToNumber_DOUBLE
                    fToNumber = CDbl(pVal)
                Case BrowsefToNumber.fToNumber_SHORT
                    fToNumber = CShort(pVal)
                Case Else
                    fToNumber = CDec(pVal)
            End Select
        End If
        Exit Function
fToNumberErr:
        xErrm = Err.Description
        Err.Raise(vbObjectError + 100, , "fToNumber:数値変換に失敗しました｡" & vbCrLf & xErrm)
    End Function

    ''' <summary>fToNumberEx -- 数値化　　※型指定省略</summary>
    ''' <param name="pVal">対象変数</param>
    ''' <returns>Decimal型の数値</returns>
    ''' <remarks>
    ''' Option Strict = True の対応
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Function fToNumberEx(ByVal pVal As Object) As Decimal

        '' Decimal型を指定して処理する
        Return fToNumberEx(Of Decimal)(pVal)

    End Function

    ''' <summary>fToNumberEx -- 数値化</summary>
    ''' <typeparam name="NumericType">返却値の数値ﾀｲﾌﾟ</typeparam>
    ''' <param name="pVal">対象変数</param>
    ''' <returns>NumericTypeで指定された型の数値</returns>
    ''' <remarks>
    ''' Option Strict = True の対応
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Function fToNumberEx(Of NumericType)(ByVal pVal As Object) As NumericType

        Dim xRet As Object

        Try

            '' 数値化できない場合は「0」で扱う
            If IsNumeric(pVal) = False Then
                xRet = 0
            Else
                xRet = pVal
            End If

            '' NumericTypeにキャストして戻す
            Return CType(xRet, NumericType)

        Catch ex As Exception

            Throw New Exception("fToNumberEx:数値変換に失敗しました｡" & vbCrLf & ex.Message)

        End Try

    End Function

    ''' <summary>fToText -- 文字化</summary>
    ''' <param name="pVal">対象変数</param>
    ''' <returns>文字列</returns>
    ''' <remarks>pValをﾃｷｽﾄ項目化します｡NULL指定された場合は空白が返却されます｡</remarks>
    ''' <作成者></作成者>
    ''' <備考>DB項目よりテキスト項目への転記など，Nullが許されない項目への転記に使用する</備考>
    <DebuggerStepThrough()>
    Public Shared Function fToText(ByVal pVal As Object) As String
        Dim ResultStr As String = String.Empty

        Try
            If pVal Is Nothing Then ''Nothing設定の場合は空白返し
                'fToText = ""
            ElseIf IsDBNull(pVal) Then
                'fToText = ""
            Else
                ResultStr = fChangeTilde(CStr(pVal))
            End If
        Catch ex As Exception
            'fToText = ""
            ResultStr = String.Empty
        End Try

        Return ResultStr

    End Function

    ''' <summary>
    ''' fTrimSingleQuotesの関数
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <history version="1" date="2017.09.21" name="PHONG-BTT">新規作成</history>
    Public Shared Function fTrimSingleQuotes(ByVal str As String) As String

        Return str.Trim("'")

    End Function

    ''' <summary>fToDate -- 日付項目化</summary>
    ''' <param name="pDate">対象変数</param>
    ''' <param name="pErrDate">エラー時の値</param>
    ''' <param name="pRetType">返却値の日付ﾀｲﾌﾟ(省略可 1:日付+時間,2:日付のみ,3:時間のみ) 既定値 1:日付+時間</param>
    ''' <returns>日付</returns>
    ''' <remarks>pDateを日付目化します｡日付項目外が指定された場合は,ﾊﾟﾗﾒｰﾀ2を,省略された場合はｼｽﾃﾑ最小値を返却します</remarks>
    ''' <作成者>米井 優顕</作成者>
    ''' <備考></備考>
    <DebuggerStepThrough()>
    Public Shared Function fToDate(ByVal pDate As Object, Optional ByVal pErrDate As Date = #1/1/1000#, Optional ByVal pRetType As BrowsefToDate = BrowsefToDate.fToDate_DateTime) As Date
        Dim xDate As Date
        Dim ResultDate As Date

        Try
            '' 日付化が可能か？
            If IsDate(pDate) Then
                xDate = CDate(pDate)
            Else
                xDate = pErrDate
            End If

            ''返却値別に型変換
            Select Case pRetType
                Case BrowsefToDate.fToDate_DateTime
                    ResultDate = xDate
                Case BrowsefToDate.fToDate_DateOnly
                    ResultDate = xDate.Date
                Case Else
                    ResultDate = CDate("#" & xDate.Hour.ToString & ":" & xDate.Minute.ToString & ":" & xDate.Millisecond & "#")
            End Select

        Catch ex As Exception

            ResultDate = pErrDate

        End Try

        Return ResultDate

    End Function

    ''' <summary>fToInt -- Int化</summary>
    ''' <param name="pVal">変換対象値</param>
    ''' <returns>Integer</returns>
    ''' <remarks>内部でfToNumberExへ処理を引き渡す</remarks>
    ''' <history version="1" date="2008.04.07" name="千葉 友則">新規作成</history> 
    <DebuggerStepThrough()>
    Public Shared Function fToInt(ByVal pVal As Object) As Integer

        Return fToNumberEx(Of Integer)(pVal)

    End Function

    ''' <summary>fToDbl -- Double化</summary>
    ''' <param name="pVal">変換対象値</param>
    ''' <returns>Integer</returns>
    ''' <remarks>内部でfToNumberExへ処理を引き渡す</remarks>
    ''' <history version="1" date="2008.04.07" name="千葉 友則">新規作成</history> 
    <DebuggerStepThrough()>
    Public Shared Function fToDbl(ByVal pVal As Object) As Double

        Return fToNumberEx(Of Double)(pVal)

    End Function

    ''' <summary>fToDec -- Decimal化</summary>
    ''' <param name="pVal">変換対象値</param>
    ''' <returns>Integer</returns>
    ''' <remarks>内部でfToNumberExへ処理を引き渡す</remarks>
    ''' <history version="1" date="2008.04.07" name="千葉 友則">新規作成</history> 
    <DebuggerStepThrough()>
    Public Shared Function fToDec(ByVal pVal As Object) As Decimal

        Return fToNumberEx(pVal)

    End Function

    '■MidB
    ''' <summary>Mid関数のバイト版。文字数と位置をバイト数で指定して文字列を切り抜く。</summary>
    ''' <param name="str">対象の文字列</param>
    ''' <param name="Start">切り抜き開始位置。全角文字を分割するよう位置が指定された場合、戻り値の文字列の先頭は意味不明の半角文字となる。</param>
    ''' <param name="Length">切り抜く文字列のバイト数</param>
    ''' <returns>切り抜かれた文字列</returns>
    ''' <remarks>最後の１バイトが全角文字の半分になる場合、その１バイトは無視される。</remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    ''' <history version="3" date="2009.06.01" name="高嶋 直樹">切り抜いた文字列のバイト数取得不正を修正</history>
    Public Shared Function MidB(ByVal str As String, ByVal Start As Integer, Optional ByVal Length As Integer = 0) As String
        '▼空文字に対しては常に空文字を返す

        If str Is Nothing Then
            Return String.Empty
        End If

        If str = "" Then
            Return ""
        End If

        '▼Lengthのチェック

        'Lengthが0か、Start以降のバイト数をオーバーする場合はStart以降の全バイトが指定されたものとみなす。

        Dim RestLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(str) - Start + 1

        If Length = 0 OrElse Length > RestLength Then
            Length = RestLength
        End If

        '▼切り抜き

        Dim SJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
        Dim B() As Byte = CType(Array.CreateInstance(GetType(Byte), Length), Byte())

        Array.Copy(SJIS.GetBytes(str), Start - 1, B, 0, Length)

        Dim st1 As String = SJIS.GetString(B)

        '▼切り抜いた結果、最後の１バイトが全角文字の半分だった場合、その半分は切り捨てる。

        ''2009.06.01 UPD START TAKASHIMA
        'Dim ResultLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(st1) - Start + 1
        Dim ResultLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(st1)
        ''2009.06.01 UPD END

        If Asc(Strings.Right(st1, 1)) = 0 Then
            'VB.NET2002,2003の場合、最後の１バイトが全角の半分の時
            Return st1.Substring(0, st1.Length - 1)
        ElseIf Length = ResultLength - 1 Then
            'VB2005の場合で最後の１バイトが全角の半分の時
            Return st1.Substring(0, st1.Length - 1)
        Else
            'その他の場合
            Return st1
        End If

    End Function

    ''' <summary>
    ''' LobDataDivision
    ''' </summary>
    ''' <param name="pLobStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    Public Shared Function LobDataDivision(ByVal pLobStr As String) As Array

        If pLobStr Is Nothing Then
            pLobStr = String.Empty
        End If

        Dim RetArr() As String
        Dim nByteLen As Decimal = LenB(pLobStr)

        '' バイト数が４０００バイトを超えるかチェック
        If nByteLen > 4000 Then
            '' 配列サイズを決定
            Dim nArrCount As Decimal = fRound((nByteLen / 4000), 1, 1) - 1
            ReDim RetArr(nArrCount)

            '' 配列に分割して値を設定
            Dim nCnt As Integer = 0
            Dim nTotalByte As Decimal = 0

            For i As Integer = nCnt To nArrCount

                '' トータルバイト数の初期化
                nTotalByte = 0

                '' 現在のトータルバイト数を取得
                For j As Integer = 0 To nArrCount
                    If RetArr(j) <> Nothing Then
                        nTotalByte = nTotalByte + fToNumberEx(Of Decimal)(LenB(RetArr(j)))
                    End If
                Next

                '' バイト分割
                If i = nArrCount Then
                    RetArr(i) = MidB(pLobStr, (nTotalByte + 1), nByteLen - nTotalByte)
                Else
                    RetArr(i) = MidB(pLobStr, (nTotalByte + 1), 4000)
                End If
            Next

        Else
            '' ４０００バイト以内の場合はそのまま返却
            ReDim RetArr(0)
            RetArr(0) = pLobStr
        End If

        Return RetArr

    End Function

    ''' <summary>Camelize -- アンダースコア繋ぎの文字列をキャメル形式・パスカル形式に変換します</summary>
    ''' <param name="str">文字列</param>
    ''' <param name="toUpper">先頭文字を大文字にするか（True: する / False: しない）</param>
    ''' <returns>String</returns>
    ''' <remarks>デフォルトはキャメル形式（先頭小文字）</remarks>
    ''' <history version="1" date="2009.08.27" name="伊藤 匡明">新規作成</history> 
    Public Shared Function Camelize(ByVal str As String, Optional ByVal toUpper As Boolean = False) As String

        Dim regex As New System.Text.RegularExpressions.Regex("[^_]+")
        Dim mc As System.Text.RegularExpressions.MatchCollection = regex.Matches(str.ToLower())
        Dim builder As StringBuilder = New StringBuilder
        For Each m As System.Text.RegularExpressions.Match In mc
            builder.Append(m.Groups(0).ToString.Substring(0, 1).ToUpper())
            builder.Append(m.Groups(0).ToString.Substring(1))
        Next
        Dim result As String = builder.ToString()
        If Not toUpper Then
            result = result.Substring(0, 1).ToLower & result.Substring(1)
        End If

        Return result

    End Function

    ''' <summary>Underscore -- キャメル形式・パスカル形式の文字列をアンダースコア繋ぎに変換します</summary>
    ''' <param name="str">文字列</param>
    ''' <param name="toUpper">大文字にするか（True: 大文字 / False: 小文字）</param>
    ''' <returns>String</returns>
    ''' <remarks>デフォルトは大文字</remarks>
    ''' <history version="1" date="2009.08.27" name="伊藤 匡明">新規作成</history> 
    Public Shared Function Underscore(ByVal str As String, Optional ByVal toUpper As Boolean = True) As String

        Dim regex As New System.Text.RegularExpressions.Regex("[A-Z]+[a-z\\d]+")
        Dim mc As System.Text.RegularExpressions.MatchCollection = regex.Matches(str)
        Dim builder As StringBuilder = New StringBuilder
        For Each m As System.Text.RegularExpressions.Match In mc
            builder.Append("_")
            builder.Append(m.Groups(0).ToString.ToUpper)
        Next
        Dim result As String = builder.ToString()
        If Not toUpper Then
            result = result.ToLower
        End If

        Return result

    End Function

#End Region

#Region " エラー通知 "

    ''' <summary>ShowExceptionMessage -- 例外からメッセージボックスを表示する</summary>
    ''' <param name="ex">例外クラス</param>
    ''' <param name="MessageTitle">メッセージタイトル</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    ''' <history version="3" date="2009.04.24" name="千葉 友則">エラーメッセージの生成を別メソッドに分離</history>
    Public Shared Sub ShowExceptionMessage(ByVal ex As Exception, ByVal MessageTitle As String)
        '' =================================================================
        ''   [MessageTitleを表示]
        ''   ---------------------------------------------------------    
        ''
        ''   【エラー内容】
        ''   [ex.Messageを表示]    
        ''
        ''   【追加情報】　　　　　　　　　　　　　　　　　'''exにInnerExceptionが設定されている場合に表示
        ''   [ex.InnerException.Messageを表示]    
        ''
        ''   【エラー内容】
        ''   [ex.StackTraceを表示]    
        '' =================================================================

        If MessageTitle Is Nothing Then
            MessageTitle = "PROCES.S5"
        End If
        ' '' 通知内容を編集
        'Dim Message As New System.Text.StringBuilder
        'Message.AppendLine("【エラー内容】")
        'Message.AppendLine(ex.Message)
        'Message.AppendLine()
        'If Not IsNothing(ex.InnerException) Then
        '    Message.AppendLine("【追加情報】")
        '    Message.AppendLine(ex.InnerException.Message)
        '    Message.AppendLine()
        'End If
        'Message.AppendLine("【スタックトレース】")
        'Message.AppendLine(ex.StackTrace)

        '' メッセージ表示
        MessageBox.Show(CreateExceptionString(ex), MessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="lev"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2009.04.24" name="千葉 友則">新規作成</history>
    Public Shared Function CreateExceptionString(ByVal ex As Exception, Optional ByVal lev As Integer = 0) As String

        Dim Message As New System.Text.StringBuilder

        If lev.Equals(0) Then
            Message.AppendLine("エラー情報")
            Message.AppendLine()
        Else
            Message.AppendLine()
            Message.AppendLine("===================================================")
            Message.AppendFormat("追加エラー情報 {0}", fToText(lev)) : Message.AppendLine()
            Message.AppendLine()
        End If

        Dim exceptionType As String
        If ex.GetType().Name = "ServiceCommonException" Then
            exceptionType = ex.ToString()
        Else
            exceptionType = ex.GetType().Name
        End If

        Message.AppendLine("【エラーの種類】")
        Message.AppendLine()
        Message.AppendLine(exceptionType)
        Message.AppendLine()
        Message.AppendLine("【内容】")
        Message.AppendLine()
        Message.AppendLine(ex.Message)
        Message.AppendLine()
        Message.AppendLine("【スタックトレース】")
        Message.AppendLine()
        Message.AppendLine(ex.StackTrace)

        If Not IsNothing(ex.InnerException) Then
            Message.AppendLine(CreateExceptionString(ex.InnerException, lev + 1))
        End If

        Return Message.ToString

    End Function

    ''' <summary>ConvertException -- 各種例外クラスを基底例外クラスに変換する</summary>
    ''' <param name="ex">例外クラス</param>
    ''' <returns>基底例外クラス</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertException(ByVal ex As Exception) As Exception

        Dim Result As Exception

        If ex.GetType Is GetType(SoapException) Then
            '' Webサービスで発生した例外クラスを基底の例外クラスに変換
            Result = New Exception(CType(ex, SoapException).Detail.ChildNodes(0).InnerText, ex)

        Else
            '' これまでの処理で変換されない例外クラスは、単純に例外クラスにキャストする
            Result = CType(ex, Exception)
        End If

        Return Result

    End Function

    ''' <summary>
    ''' WriteEventLog -- イベントログを出力する
    ''' </summary>
    ''' <param name="pEx">例外クラス</param>
    ''' <param name="pMsg">エラーメッセージ</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    ''' <history version="3" date="2011.12.14" name="内田 久義">(SCで落ちる原因になっていた為)EventLogへの書き込みエラーは何もしようがないのでスルーする</history>
    Public Shared Sub WriteEventLog(ByVal pEx As Exception, ByVal pMsg As String)

        If pMsg Is Nothing Then
            pMsg = String.Empty
        End If

        Dim xMsg As String = String.Empty

        '' 2010.01.26 UPD START
        'xMsg = pMsg & System.Environment.NewLine & "[StackTrace:]" & pEx.GetType.ToString & _
        '        System.Environment.NewLine & pEx.StackTrace
        xMsg &= pMsg
        xMsg &= System.Environment.NewLine
        xMsg &= CreateExceptionString(pEx)
        '' 2010.01.26 UPD END

        ''2011.12.14 UPD START
        'EventLog.WriteEntry("PROCES.S5", xMsg, EventLogEntryType.Error, 50000)
        Try
            EventLog.WriteEntry("PROCES.S5", xMsg, EventLogEntryType.Error, 50000)
        Catch dummyex As Exception
            ''(SCで落ちる原因になっていた為)EventLogへの書き込みエラーは何もしようがないのでスルーする
        End Try
        ''2011.12.14 UPD END

    End Sub

#End Region

#Region " 別アプリ起動処理 "

    ' ''' <summary>
    ' ''' ClicOnce アプリケーションを起動するための URL を作成します。
    ' ''' </summary>
    ' ''' <param name="deploymentUrl"></param>
    ' ''' <param name="kidoName"></param>
    ' ''' <param name="query"></param>
    ' ''' <returns></returns>
    'Public Shared Function CreateApplicationDownloadUrl(deploymentUrl As String, kidoName As String, query As StartUpQuery) As String

    '    Dim queryString = query.ToQueryString()

    '    ' hoge.application or hoge => hoge
    '    Dim applicationName = System.IO.Path.ChangeExtension(kidoName, Nothing)

    '    Return $"{deploymentUrl.TrimEnd("/"c)}/{applicationName}/{applicationName}.application?{queryString}"

    'End Function

    ' ''' <summary>
    ' ''' ClicOnce アプリケーションを起動するための IE のプロセス情報を作成します。
    ' ''' </summary>
    ' ''' <param name="deploymentUrl"></param>
    ' ''' <param name="kidoName"></param>
    ' ''' <param name="query"></param>
    ' ''' <returns></returns>
    'Public Shared Function CreateProcessInfo(deploymentUrl As String, kidoName As String, query As StartUpQuery) As ProcessStartInfo

    '    Return CreateProcessInfo(CreateApplicationDownloadUrl(deploymentUrl, kidoName, query), ProcessWindowStyle.Hidden)

    'End Function

    ''' <summary>CretateProcessInfo -- </summary>
    ''' <param name="arg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2009.07.22" name="千葉 友則">オーバーロード実装（引数をそのままセットする）</history>
    Public Shared Function CreateProcessInfo(ByVal arg As String, ByVal flg As ProcessWindowStyle) As System.Diagnostics.ProcessStartInfo

        Dim info As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo

        ''起動情報の作成（毎回新しいIEを介して起動するように変更）
        info.FileName = "IExplore.exe"                                      '' IEプロセス名（固定）
        info.Arguments = arg                                                '' 引数をダイレクトにセット
        info.UseShellExecute = True                                         '' シェル起動
        info.CreateNoWindow = False                                         '' 新しいWindowを作成
        info.WindowStyle = flg                                              '' 引数をダイレクトにセット

        Return info

    End Function

#End Region

#Region " 状態ラベル変更処理 "

    ''' <summary>
    ''' SetLabelConfig
    ''' </summary>
    ''' <param name="LabelControl"></param>
    ''' <param name="Index"></param>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">初版不明</history>
    ''' <history version="2" date="2009.03.25" name="千葉 友則">引数判定追加</history>
    Public Shared Sub SetLabelConfig(ByVal LabelControl As Label, ByVal Index As BrouseConfig)

        If LabelControl Is Nothing Then
            Return
        End If

        Select Case Index
            Case BrouseConfig.NoData
                '未登録文字
                LabelControl.Text = "コード未入力"
                LabelControl.ForeColor = System.Drawing.Color.Khaki
                LabelControl.BackColor = System.Drawing.Color.Olive
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.NewData
                '新規登録文字(入力開始前)
                LabelControl.Text = "未登録コード"
                LabelControl.ForeColor = System.Drawing.Color.PowderBlue
                LabelControl.BackColor = System.Drawing.Color.Navy
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.NewDataInput
                '新規登録文字(入力開始後)
                LabelControl.Text = "新規入力中"
                LabelControl.ForeColor = System.Drawing.Color.Navy
                LabelControl.BackColor = System.Drawing.Color.PowderBlue
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.UpdData
                '修正登録文字(入力開始前)
                LabelControl.Text = "照会表示中"
                LabelControl.ForeColor = System.Drawing.Color.PaleGreen
                LabelControl.BackColor = System.Drawing.Color.DarkGreen
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.UpdDataInput
                '修正登録文字(入力開始後)
                LabelControl.Text = "修正入力中"
                LabelControl.ForeColor = System.Drawing.Color.DarkGreen
                LabelControl.BackColor = System.Drawing.Color.PaleGreen
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.NoSec
                '権限無し文字
                LabelControl.Text = "未権限コード"
                LabelControl.ForeColor = System.Drawing.Color.Red
                LabelControl.BackColor = System.Drawing.Color.Black
                LabelControl.BorderStyle = BorderStyle.FixedSingle
                ''2008.07.01 ADD START T.Nara
            Case BrouseConfig.Running
                'データ更新実行中文字
                LabelControl.Text = "更新中"
                LabelControl.ForeColor = System.Drawing.Color.LightPink
                LabelControl.BackColor = System.Drawing.Color.Maroon
                LabelControl.BorderStyle = BorderStyle.FixedSingle
                ''2008.07.01 ADD END
            Case BrouseConfig.Initialize
                '初期化(ラベルの非表示)
                LabelControl.Text = ""
                LabelControl.ForeColor = LabelControl.Parent.BackColor
                LabelControl.BackColor = LabelControl.Parent.BackColor
                LabelControl.BorderStyle = BorderStyle.None
        End Select

    End Sub

#End Region

#Region " 部門権限関連 "

#Region " 部門権限 "
    ''' <summary>
    ''' GetWhereStrBmnSecurity -- 部門権限を考慮するためのWhere句を取得
    ''' </summary>
    ''' <param name="checkType">CheckBmnSecType列挙体</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where句</returns>
    ''' <remarks>SearchPanelからの呼出も考慮してCommonMethodに実装
    '''          Formアプリで使用する場合はBaseFormの同名Methodを呼出すこと</remarks>
    ''' <history version="1" date="2010.05.13" name="清水 勝也">新規作成</history>
    Public Overloads Shared Function GetWhereStrBmnSecurity(ByVal checkType As CheckBmnSecType,
                                                            Optional ByVal userNo As Integer = 0) As String
        ' 通常はCODEで権限チェック
        Return GetWhereStrBmnSecurity("CODE", checkType, userNo)
    End Function

    ''' <summary>
    ''' GetWhereStrBmnSecurity -- 部門権限を考慮するためのWhere句を取得
    ''' </summary>
    ''' <param name="field">Field名 in (権限マスタ.CODE)</param>
    ''' <param name="checkType">CheckBmnSecType列挙体</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where句</returns>
    ''' <remarks>SearchPanelからの呼出も考慮してCommonMethodに実装
    '''          Formアプリで使用する場合はBaseFormの同名Methodを呼出すこと</remarks>
    ''' <history version="1" date="2010.05.13" name="清水 勝也">新規作成</history>
    Public Overloads Shared Function GetWhereStrBmnSecurity(ByVal field As String,
                                                            ByVal checkType As CheckBmnSecType,
                                                            Optional ByVal userNo As Integer = 0) As String
        Dim xSqlTxt As String

        ' Where句のSQL生成
        xSqlTxt = String.Format("		{0} IN ( ", field)  ' 指定されたFieldを使用
        xSqlTxt &= vbCrLf & "				SELECT "
        xSqlTxt &= vbCrLf & "						CODE "
        xSqlTxt &= vbCrLf & "				FROM "
        xSqlTxt &= vbCrLf & "						{0} "
        xSqlTxt &= vbCrLf & "				WHERE "
        xSqlTxt &= vbCrLf & "						USER_NO	=	{1} "
        xSqlTxt &= vbCrLf & "		) "

        ' 専用の変換処理をして返却
        Return FormatStrBmnSecurity(xSqlTxt, checkType, userNo)
    End Function
#End Region
#Region " 部門階層権限 "
    ''' <summary>
    ''' GetWhereStrBmnKaisouSecurity -- 部門権限を考慮するためのWhere句を取得
    ''' </summary>
    ''' <param name="checkType">CheckBmnSecType列挙体</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where句</returns>
    ''' <remarks>SearchPanelからの呼出も考慮してCommonMethodに実装
    '''          Formアプリで使用する場合はBaseFormの同名Methodを呼出すこと</remarks>
    ''' <history version="1" date="2010.05.13" name="清水 勝也">新規作成</history>
    Public Overloads Shared Function GetWhereStrBmnKaisouSecurity(ByVal checkType As CheckBmnSecType,
                                                                  Optional ByVal userNo As Integer = 0) As String
        ' 通常はCODEで権限チェック
        Return GetWhereStrBmnKaisouSecurity("CODE", checkType, userNo)
    End Function

    ''' <summary>
    ''' GetWhereStrBmnSecurity -- 部門権限を考慮するためのWhere句を取得
    ''' </summary>
    ''' <param name="field">Field名 in (権限マスタ.CODE)</param>
    ''' <param name="checkType">CheckBmnSecType列挙体</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where句</returns>
    ''' <remarks>SearchPanelからの呼出も考慮してCommonMethodに実装
    '''          Formアプリで使用する場合はBaseFormの同名Methodを呼出すこと</remarks>
    ''' <history version="1" date="2010.05.13" name="清水 勝也">新規作成</history>
    Public Overloads Shared Function GetWhereStrBmnKaisouSecurity(ByVal field As String,
                                                                  ByVal checkType As CheckBmnSecType,
                                                                  Optional ByVal userNo As Integer = 0) As String
        Dim xSqlTxt As String

        ' Where句のSQL生成
        xSqlTxt = String.Format("		{0} IN ( ", field)  ' 指定されたFieldを使用
        xSqlTxt &= vbCrLf & "				SELECT "
        xSqlTxt &= vbCrLf & "						KAISOU "
        xSqlTxt &= vbCrLf & "				FROM "
        xSqlTxt &= vbCrLf & "						{0} "
        xSqlTxt &= vbCrLf & "				WHERE "
        xSqlTxt &= vbCrLf & "						USER_NO	=	{1} "
        xSqlTxt &= vbCrLf & "				GROUP BY "
        xSqlTxt &= vbCrLf & "						KAISOU "
        xSqlTxt &= vbCrLf & "		) "

        ' 専用の変換処理をして返却
        Return FormatStrBmnSecurity(xSqlTxt, checkType, userNo)
    End Function
#End Region

    ''' <summary>
    ''' FormatStrBmnSecurity(Private) -- 部門権限考慮用のString.Format
    ''' </summary>
    ''' <param name="xSqlTxt">変換項目を含むString({0}:権限マスタ, {1}:USER_NO)</param>
    ''' <param name="checkType">CheckBmnSecType列挙体</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>変換後のString</returns>
    ''' <remarks>Private</remarks>
    ''' <history version="1" date="2010.05.13" name="清水 勝也">新規作成</history>
    Private Shared Function FormatStrBmnSecurity(ByVal xSqlTxt As String, ByVal checkType As CheckBmnSecType, ByVal userNo As Integer) As String

        Dim xBmnSecMaster As String
        Select Case checkType
            Case CheckBmnSecType.Sansyo
                xBmnSecMaster = "S_BMN_UW" ' 参照権限
            Case CheckBmnSecType.Input
                xBmnSecMaster = "I_BMN_UW" ' 入力権限
            Case Else
                xBmnSecMaster = "S_BMN_UW" ' 参照権限(Else)
        End Select

        Dim oUserNo As Object
        If userNo = 0 Then
            oUserNo = "@SYSUSERNO@" ' ServiceCommonで置換されるUSER_NOの代替Text
        Else
            oUserNo = fSqlStr(userNo, BrowsefSqlStr.fSqlStr_Integer)
        End If

        ' 変換して返却
        Return String.Format(xSqlTxt, xBmnSecMaster, oUserNo)
    End Function

#End Region

End Class



