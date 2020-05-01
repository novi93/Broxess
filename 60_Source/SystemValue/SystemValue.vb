
#Region "imports宣言"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region


''' <summary>
''' システム変数
''' </summary>
''' <remarks></remarks>
''' <history version="5.0.0.0" date="2008/04/07" name="宮本 淳也">新規作成</history> 
''' <history version="5.0.0.0" date="2008/04/08" name="千葉 友則">印刷指示で使用する初期値設定項目を追加</history>
''' <history version="5.0.0.0" date="2008/04/08-2" name="千葉 友則">SysUserNo追加</history>
''' <history version="5.0.0.0" date="2008/04/22" name="石寺 健士">SysMindateを1000/01/01に修正、SysDefCodeN,SysDefCodeXを追加</history>
''' <history version="5.0.0.0" date="2008/05/01" name="石寺 健士">文字コード数値運用時桁数(SysCharKeta)の追加</history>
''' <history version="5.0.0.0" date="2008.07.16" name="千葉 友則">WF用分離ファイルを生成</history>
''' <history version="5.0.0.0" date="2009.02.25" name="石澤 雪乃">部門コードの文字列化対応</history>
''' <history version="5.0.0.0" date="2009.03.28" name="内田 久義">SysMinCode,SysMaxCodeの文字型SysMinCodeX,SysMaxCodeXを追加</history>
''' <history version="5.0.0.0" date="2009.05.27" name="内田 久義">SysBmnKaisouの入力用(SysBmnKaisouIn)を追加</history>
''' <history version="5.0.0.0" date="2009.11.05" name="宮本 淳也">現場用共通変数の追加</history>
''' <history version="5.0.0.0" date="2010.03.05" name="宮本 淳也">現場用共通変数の基本予算番号の追加</history>
''' <history version="5.0.1.1" date="2011.12.15" name="内田 久義">アセンブリ情報・参照設定を正しく設定し直し</history>
''' <history version="5.0.2.0" date="2012.09.15" name="内田 久義">ダミーのコンストラクタを追加</history>
Public Class SystemValue
#Region "プロパティ"
    Public Shared SysOpid As String = String.Empty
    Public Shared SysUserNo As Integer = 0
    Public Shared SysKais As Integer = 0
    Public Shared SysDDate As Date = #12/31/2087#
    Public Shared SysLoginTime As Date = #12/31/2087#
    Public Shared SysLogoutTime As Date = #12/31/2087#
    Public Shared SysMenuPId As Integer = 0
    Public Shared SysDnoStart As Integer = 0
    Public Shared SysDnoEnd As Integer = 0
    Public Shared SysTant As Integer = 0

    ''2009.02.25 UPD START ISHIZAWA 部門コードの文字コード化対応
    Public Shared SysBmnUw As String = String.Empty
    Public Shared SysBmnUwIn As String = String.Empty

    'Public Shared SysBmnUw As Integer = 0
    'Public Shared SysBmnUwIn As Integer = 0
    'Public Shared SysBmn As Integer = 0
    ''2009.02.25 UPD END
    Public Shared SysBmnKaisou As Integer = 0
    ''2009.05.27 ADD START
    Public Shared SysBmnKaisouIn As Integer = 0
    ''2009.05.27 ADD END

    Public Shared SysMinCode As Integer = 0
    Public Shared SysMaxCode As Integer = 999999999
    Public Shared SysMinDate As Date = #1/1/1000#
    Public Shared SysMaxDate As Date = #12/31/2087#

    Public Shared ReadOnly SysDefCodeN As Integer = -1
    Public Shared ReadOnly SysDefCodeX As String = " "
    ''==========================================================
    '' 管理マスタ　CODE=1 アイテム
    Public Shared SysDateFormat As Integer = 0

    Public Shared SysDatePrint As Integer = 0
    Public Shared SysKaisPrint As Integer = 0
    Public Shared SysReportPrint As Integer = 0
    ''==========================================================
    '' 管理マスタ　CODE=2 アイテム
    Public Shared SysSKojFormatType As Integer = 0
    Public Shared SysKojFormatType As Integer = 0
    Public Shared SysKojUwFormatType As Integer = 0
    Public Shared SysTorFormatType As Integer = 0

    Public Shared SysSKojFormat As String = String.Empty
    Public Shared SysKojFormat As String = String.Empty
    Public Shared SysKojUwFormat As String = String.Empty
    Public Shared SysTorFormat As String = String.Empty

    Public Shared SysCharKeta As Integer = 10

    ''2009.03.28 ADD START
    Public Shared SysMinCodeX As String = "        0"
    Public Shared SysMaxCodeX As String = "999999999"
    ''2009.03.28 ADD END

    Public Shared SysMinSKoj As String = String.Empty
    Public Shared SysMaxSKoj As String = "ZZZZZZZZZ"
    Public Shared SysMinKoj As String = String.Empty
    Public Shared SysMaxKoj As String = "ZZZZZZZZZ"
    Public Shared SysMinKojUw As String = String.Empty
    Public Shared SysMaxKojUw As String = "ZZZZZZZZZ"
    Public Shared SysMinTor As String = String.Empty
    Public Shared SysMaxTor As String = "ZZZZZZZZZ"
    Public Shared SysKojUwKihon As String = String.Empty
    ''==========================================================
    ''2009.02.25 ADD START ISHIZAWA 部門コードの文字コード化対応
    '' 管理マスタ　CODE=4 アイテム
    Public Shared SysBmnFormatType As Integer = 0
    Public Shared SysBmnFormat As String = String.Empty
    Public Shared SysMinBmn As String = String.Empty
    Public Shared SysMaxBmn As String = "ZZZZZZZZZ"
    ''2009.02.25 ADD END

#End Region

#Region " 内部変数"
    Private Ret As Boolean              ' 判定用
#End Region

#Region "コンストラクタ"

    ''2012.09.15 ADD START
    ''' <summary>
    ''' New
    ''' </summary>
    ''' <remarks>ダミー</remarks>
    ''' <history version="1" date="2012.09.15" name="内田 久義">新規作成</history>
    Public Sub New()

    End Sub
    ''2012.09.15 ADD END

    ''' <summary>
    ''' New
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2008/04/07" name="宮本 淳也">新規作成</history>
    Public Sub New(ByVal dr As DataRow)

        ''配列に取得値を設定します。
        'SetSystemSetting(dr)

    End Sub

#End Region

    '#Region "メソッド"
    '    ''' <summary>
    '    ''' 変数値の展開
    '    ''' </summary>
    '    ''' <param name="dr">SystemLoadで取得したシステムデータ列</param>
    '    ''' <remarks></remarks>
    '    ''' <history version="1" date="2008/04/07" name="宮本 淳也">新規作成</history> 
    '    ''' <history version="2" date="2008/04/08" name="千葉 友則">SysUserNoを追加</history>
    '    ''' <history version="3" date="2008/04/22" name="石寺 健士">SysDefCodeN,SysDefCodeXを追加</history> 
    '    ''' <history version="4" date="2008/05/01" name="石寺 健士">文字コード数値運用時桁数(SysCharKeta)の追加</history>
    '    ''' <history version="5" date="2008/06/23" name="奈良 孝志">SysDDateの時刻部分除外</history>
    '    ''' <history version="6" date="2008.07.16" name="千葉 友則">WF設定用の呼び出しメソッドを追加</history>
    '    ''' <history version="7" date="2009.01.07" name="中山 俊史">販売管理設定用の呼び出しメソッドを追加</history>
    '    ''' <history version="8" date="2009.02.25" name="石澤 雪乃">部門コードの文字コード化対応</history>
    '    ''' <history version="9" date="2009.03.28" name="内田 久義">SysMinCode,SysMaxCodeの文字型SysMinCodeX,SysMaxCodeXを追加</history>
    '    ''' <history version="10" date="2009.05.27" name="内田 久義">SysBmnKaisouの入力用(SysBmnKaisouIn)を追加</history>
    '    ''' <history version="11" date="2009.11.05" name="宮本 淳也">現場設定用の呼び出しメソッドを追加</history>
    '    Public Shared Sub SetSystemSetting(ByVal dr As DataRow)

    '        SysOpid = fToText(dr("SysOpid"))
    '        SysUserNo = fToInt(dr("SysUserNo"))
    '        SysKais = fToInt(dr("SysKais"))
    '        ''2008.06.23 UPD START T.Nara 時刻部分除外
    '        'SysDDate = fToDate(dr("SysDDate"))
    '        SysDDate = fToDate(dr("SysDDate"), #1/1/1000#, BrowsefToDate.fToDate_DateOnly)
    '        ''2008.06.23 UPD END
    '        SysLoginTime = fToDate(dr("SysLoginTime"))
    '        SysLogoutTime = fToDate(dr("SysLogoutTime"))
    '        SysMenuPId = fToInt(dr("SysMenuPId"))
    '        SysDnoStart = fToInt(dr("SysDnoStart"))
    '        SysDnoEnd = fToInt(dr("SysDnoEnd"))
    '        SysTant = fToInt(dr("SysTant"))

    '        ''2009.02.25 UPD START ISHIZAWA 部門コードの文字コード化対応
    '        SysBmnUw = fToText(dr("SysBmnUw"))
    '        SysBmnUwIn = fToText(dr("SysBmnUwIn"))
    '        'SysBmnUw = fToInt(dr("SysBmnUw"))
    '        'SysBmnUwIn = fToInt(dr("SysBmnUwIn"))
    '        'SysBmn = fToInt(dr("SysBmn"))
    '        ''2009.02.25 UPD END
    '        SysBmnKaisou = fToInt(dr("SysBmnKaisou"))
    '        ''2009.05.27 ADD START
    '        SysBmnKaisouIn = fToInt(dr("SysBmnKaisouIn"))
    '        ''2009.05.27 ADD END

    '        SysSKojFormatType = fToInt(dr("SysSKojFormatType"))
    '        SysKojFormatType = fToInt(dr("SysKojFormatType"))
    '        SysKojUwFormatType = fToInt(dr("SysKojUwFormatType"))
    '        SysTorFormatType = fToInt(dr("SysTorFormatType"))
    '        ''2009.02.25 ADD START ISHIZAWA 部門コードの文字コード化対応
    '        SysBmnFormatType = fToInt(dr("SysBmnFormatType"))
    '        ''2009.02.25 ADD END

    '        SysCharKeta = fToInt(dr("SysCharKeta"))

    '        SysSKojFormat = fToText(dr("SysSKojFormat"))
    '        SysKojFormat = fToText(dr("SysKojFormat"))
    '        SysKojUwFormat = fToText(dr("SysKojUwFormat"))
    '        SysTorFormat = fToText(dr("SysTorFormat"))
    '        ''2009.02.25 ADD START ISHIZAWA 部門コードの文字コード化対応
    '        SysBmnFormat = fToText(dr("SysBmnFormat"))
    '        ''2009.02.25 ADD END

    '        SysMinCode = fToInt(dr("SysMinCode"))
    '        SysMaxCode = fToInt(dr("SysMaxCode"))
    '        SysMinDate = fToDate(dr("SysMinDate"))
    '        SysMaxDate = fToDate(dr("SysMaxDate"))

    '        ''2009.03.28 ADD START
    '        SysMinCodeX = fToText(dr("SysMinCodeX"))
    '        SysMaxCodeX = fToText(dr("SysMaxCodeX"))
    '        ''2009.03.28 ADD END

    '        SysMinSKoj = fToText(dr("SysMinSKoj"))
    '        SysMaxSKoj = fToText(dr("SysMaxSKoj"))
    '        SysMinKoj = fToText(dr("SysMinKoj"))
    '        SysMaxKoj = fToText(dr("SysMaxKoj"))
    '        SysMinKojUw = fToText(dr("SysMinKojUw"))
    '        SysMaxKojUw = fToText(dr("SysMaxKojUw"))
    '        SysMinTor = fToText(dr("SysMinTor"))
    '        SysMaxTor = fToText(dr("SysMaxTor"))
    '        ''2009.02.25 ADD START ISHIZAWA 部門コードの文字コード化対応
    '        SysMinBmn = fToText(dr("SysMinBmn"))
    '        SysMaxBmn = fToText(dr("SysMaxBmn"))
    '        ''2009.02.25 ADD END
    '        SysKojUwKihon = fToText(dr("SysKojUwKihon"))

    '        SysDateFormat = fToInt(dr("SysDateFormat"))
    '        SysDatePrint = fToInt(dr("SysDatePrint"))
    '        SysKaisPrint = fToInt(dr("SysKaisPrint"))
    '        SysReportPrint = fToInt(dr("SysReportPrint"))

    '        ' '' WFで使用する設定をセット
    '        'SetSystemSettingWF(dr)

    '        ' ''2009.01.07 ADD START NAKAYAMA
    '        ' '' 販売で使用する設定をセット
    '        'SetSystemSettingHanbai(dr)
    '        ' ''2009.01.07 ADD END

    '        ' ''2009.11.05 ADD START MIYAMOTO
    '        'SetSystemSettingGenba(dr)
    '        ' ''2009.11.05 ADD END
    '    End Sub

    '#End Region

End Class
