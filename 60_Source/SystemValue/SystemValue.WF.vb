
#Region "imports宣言"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region

'' SystemValueクラス分離ファイル　WF専用 
Partial Class SystemValue
#Region "プロパティ"

    Public Shared SysValueGet As Boolean = False
    'Public Shared SysLoginDate As Date = #12/31/2087#

    Public Shared SysWFSetPGCode As Integer = 209002

#End Region

#Region "メソッド"
    ''' <summary>SetSystemSettingWF -- WF用変数値の展開</summary>
    ''' <param name="dr">SystemLoadで取得したシステムデータ列</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2008.07.16" name="千葉 友則">N社SetSystemSettingより、WF部分のみ分離したメソッド</history>
    ''' <history version="2" date="2008.07.29" name="千葉 友則">WF設定のプログラムコードを追加</history>
    Public Shared Sub SetSystemSettingWF(ByVal dr As DataRow)

        'SysLoginDate = fToDate(dr("SysLoginDate"))

        'SysWFSetPGCode = fToInt(dr("SysWFSetPGCode"))

    End Sub

#End Region

End Class
