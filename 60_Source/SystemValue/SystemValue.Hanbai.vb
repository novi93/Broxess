
#Region "imports�錾"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region

'' SystemValue�N���X�����t�@�C���@�̔��Ǘ���p 
Partial Class SystemValue
#Region "�v���p�e�B"

    ''2009.03.20 ADD START NAKAYAMA
    '' �����R�[�h�������̃\�[�g��(1:�R�[�h���A2:�J�i��)
    Public Shared SysHanSortTor As Integer = 0
    '' �����R�[�h�ȊO�������̃\�[�g��(1:�R�[�h���A2:�J�i��)
    Public Shared SysHanSort As Integer = 0
    ''2009.03.20 ADD END
    Public Shared SysSokoFormatType As Integer = 0
    Public Shared SysSyohinFormatType As Integer = 0
    Public Shared SysJyuFormatType As Integer = 0
    Public Shared SysHacFormatType As Integer = 0
    Public Shared SysSyuFormatType As Integer = 0
    Public Shared SysUriFormatType As Integer = 0
    Public Shared SysNyuFormatType As Integer = 0
    Public Shared SysSirFormatType As Integer = 0

    Public Shared SysSokoFormat As String = String.Empty
    Public Shared SysSyohinFormat As String = String.Empty
    Public Shared SysJyuFormat As String = String.Empty
    Public Shared SysHacFormat As String = String.Empty
    Public Shared SysSyuFormat As String = String.Empty
    Public Shared SysUriFormat As String = String.Empty
    Public Shared SysNyuFormat As String = String.Empty
    Public Shared SysSirFormat As String = String.Empty
    Public Shared SysMinSoko As String = String.Empty
    Public Shared SysMaxSoko As String = "ZZZZZZZZZ"
    Public Shared SysMinSyohin As String = String.Empty
    Public Shared SysMaxSyohin As String = "ZZZZZZZZZ"
    Public Shared SysMinJyu As String = String.Empty
    Public Shared SysMaxJyu As String = "ZZZZZZZZZ"
    Public Shared SysMinHac As String = String.Empty
    Public Shared SysMaxHac As String = "ZZZZZZZZZ"
    Public Shared SysMinSyu As String = String.Empty
    Public Shared SysMaxSyu As String = "ZZZZZZZZZ"
    Public Shared SysMinUri As String = String.Empty
    Public Shared SysMaxUri As String = "ZZZZZZZZZ"
    Public Shared SysMinNyu As String = String.Empty
    Public Shared SysMaxNyu As String = "ZZZZZZZZZ"
    Public Shared SysMinSir As String = String.Empty
    Public Shared SysMaxSir As String = "ZZZZZZZZZ"

#End Region

#Region "���\�b�h"

    ''' <summary>SetSystemSettingHanbai -- �̔��Ǘ��p�ϐ��l�̓W�J</summary>
    ''' <param name="dr">SystemLoad�Ŏ擾�����V�X�e���f�[�^��</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2009.01.07" name="���R �r�j">SetSystemSetting���A�̔��Ǘ������̂ݕ����������\�b�h(�������t�@�C���Ɉڍs���܂�)</history>
    ''' <history version="2" date="2009.03.20" name="���R �r�j">������ʂ̃\�[�g�����Ǘ�����ϐ���ǉ�</history>
    Public Shared Sub SetSystemSettingHanbai(ByVal dr As DataRow)

        ' ''2009.03.20 ADD START NAKAYAMA
        'SysHanSortTor = fToInt(dr("SysHanSortTor"))
        'SysHanSort = fToInt(dr("SysHanSort"))
        ' ''2009.03.20 ADD END

        'SysSokoFormatType = fToInt(dr("SysSokoFormatType"))
        'SysSyohinFormatType = fToInt(dr("SysSyohinFormatType"))
        'SysJyuFormatType = fToInt(dr("SysJyuFormatType"))
        'SysHacFormatType = fToInt(dr("SysHacFormatType"))
        'SysSyuFormatType = fToInt(dr("SysSyuFormatType"))
        'SysUriFormatType = fToInt(dr("SysUriFormatType"))
        'SysNyuFormatType = fToInt(dr("SysNyuFormatType"))
        'SysSirFormatType = fToInt(dr("SysSirFormatType"))

        'SysSokoFormat = fToText(dr("SysSokoFormat"))
        'SysSyohinFormat = fToText(dr("SysSyohinFormat"))
        'SysJyuFormat = fToText(dr("SysJyuFormat"))
        'SysHacFormat = fToText(dr("SysHacFormat"))
        'SysSyuFormat = fToText(dr("SysSyuFormat"))
        'SysUriFormat = fToText(dr("SysUriFormat"))
        'SysNyuFormat = fToText(dr("SysNyuFormat"))
        'SysSirFormat = fToText(dr("SysSirFormat"))
        'SysMinSoko = fToText(dr("SysMinSoko"))
        'SysMaxSoko = fToText(dr("SysMaxSoko"))
        'SysMinSyohin = fToText(dr("SysMinSyohin"))
        'SysMaxSyohin = fToText(dr("SysMaxSyohin"))
        'SysMinJyu = fToText(dr("SysMinJyu"))
        'SysMaxJyu = fToText(dr("SysMaxJyu"))
        'SysMinHac = fToText(dr("SysMinHac"))
        'SysMaxHac = fToText(dr("SysMaxHac"))
        'SysMinSyu = fToText(dr("SysMinSyu"))
        'SysMaxSyu = fToText(dr("SysMaxSyu"))
        'SysMinUri = fToText(dr("SysMinUri"))
        'SysMaxUri = fToText(dr("SysMaxUri"))
        'SysMinNyu = fToText(dr("SysMinNyu"))
        'SysMaxNyu = fToText(dr("SysMaxNyu"))
        'SysMinSir = fToText(dr("SysMinSir"))
        'SysMaxSir = fToText(dr("SysMaxSir"))

    End Sub

#End Region

End Class
