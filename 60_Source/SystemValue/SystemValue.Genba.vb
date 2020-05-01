
#Region "imports�錾"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region

'' SystemValue�N���X�����t�@�C���@����Ǘ���p 
Partial Class SystemValue
#Region "�v���p�e�B"

    Public Shared SysSekkeiFormatType As Integer = 0

    Public Shared SysSekkeiFormat As String = String.Empty
    Public Shared SysMinSekkei As String = String.Empty
    Public Shared SysMaxSekkei As String = "ZZZZZZZZZ"
    ''2010.03.05 ADD START MIYAMOTO
    Public Shared SysSekkeiKihon As String = String.Empty
    ''2010.03.05 ADD END

#End Region

#Region "���\�b�h"

    ''' <summary>SetSystemSettingHanbai -- ����p�ϐ��l�̓W�J</summary>
    ''' <param name="dr">SystemLoad�Ŏ擾�����V�X�e���f�[�^��</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2009.11.05" name="�{�{ �~��">SetSystemSetting���A���ꕔ���̂ݕ����������\�b�h(�������t�@�C���Ɉڍs���܂�)</history>
    Public Shared Sub SetSystemSettingGenba(ByVal dr As DataRow)

        ''�Ǘ�Ͻ��̒l���Q�Ƃ���Web����̃f�[�^���n���ɂ��Ă͌�ɑΉ�
        'SysSekkeiFormatType = fToInt(dr("SysSekkeiFormatType"))
        'SysSekkeiFormat = fToText(dr("SysSekkeiFormat"))
        'SysMinSekkei = fToText(dr("SysMinSekkei"))
        'SysMaxSekkei = fToText(dr("SysMaxSekkei"))
        'SysSekkeiKihon = fToText(dr("SysSekkeiKihon"))
        ''2010.03.05 UPD START MIYAMOTO
        ''�b��l�ݒ�
        'SysSekkeiFormatType = 1
        'SysSekkeiFormat = "000-00"
        'SysMinSekkei = "000-00"
        'SysMaxSekkei = "999-99"
        ' ''2010.03.05 ADD START MIYAMOTO
        'SysSekkeiKihon = "001-00"
        ' ''2010.03.05 ADD END
        ''����V�X�e���p�����[�^�̔��肪�s���Ȃ��ꍇ�͎b��̍��ڂցA����ȊO��Web���擾����
        If dr.Table.Columns.Contains("SysSekkeiFormatType") = False Then
            '�b��l�ݒ�
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
