
#Region "imports宣言"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region

'' SystemValueクラス分離ファイル　現場管理専用 
Partial Class SystemValue
#Region "プロパティ"

    Public Shared SysSekkeiFormatType As Integer = 0

    Public Shared SysSekkeiFormat As String = String.Empty
    Public Shared SysMinSekkei As String = String.Empty
    Public Shared SysMaxSekkei As String = "ZZZZZZZZZ"
    ''2010.03.05 ADD START MIYAMOTO
    Public Shared SysSekkeiKihon As String = String.Empty
    ''2010.03.05 ADD END

#End Region

#Region "メソッド"

    ''' <summary>SetSystemSettingHanbai -- 現場用変数値の展開</summary>
    ''' <param name="dr">SystemLoadで取得したシステムデータ列</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2009.11.05" name="宮本 淳也">SetSystemSettingより、現場部分のみ分離したメソッド(※分離ファイルに移行します)</history>
    Public Shared Sub SetSystemSettingGenba(ByVal dr As DataRow)

        ''管理ﾏｽﾀの値を参照したWebからのデータ引渡しについては後に対応
        'SysSekkeiFormatType = fToInt(dr("SysSekkeiFormatType"))
        'SysSekkeiFormat = fToText(dr("SysSekkeiFormat"))
        'SysMinSekkei = fToText(dr("SysMinSekkei"))
        'SysMaxSekkei = fToText(dr("SysMaxSekkei"))
        'SysSekkeiKihon = fToText(dr("SysSekkeiKihon"))
        ''2010.03.05 UPD START MIYAMOTO
        ''暫定値設定
        'SysSekkeiFormatType = 1
        'SysSekkeiFormat = "000-00"
        'SysMinSekkei = "000-00"
        'SysMaxSekkei = "999-99"
        ' ''2010.03.05 ADD START MIYAMOTO
        'SysSekkeiKihon = "001-00"
        ' ''2010.03.05 ADD END
        ''現場システムパラメータの判定が行えない場合は暫定の項目へ、それ以外はWebより取得する
        If dr.Table.Columns.Contains("SysSekkeiFormatType") = False Then
            '暫定値設定
            SysSekkeiFormatType = 1
            SysSekkeiFormat = "000-00"
            SysMinSekkei = "000-00"
            SysMaxSekkei = "999-99"
            SysSekkeiKihon = "001-00"
        Else
            'SysSekkeiFormatType = fToInt(dr("SysSekkeiFormatType"))
            'SysSekkeiFormat = fToText(dr("SysSekkeiFormat"))
            'SysMinSekkei = fToText(dr("SysMinSekkei"))
            'SysMaxSekkei = fToText(dr("SysMaxSekkei"))
            'SysSekkeiKihon = fToText(dr("SysSekkeiKihon"))
        End If
        ''2010.03.05 UPD END
    End Sub

#End Region

End Class
