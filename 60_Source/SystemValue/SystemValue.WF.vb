
#Region "imports�錾"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region

'' SystemValue�N���X�����t�@�C���@WF��p 
Partial Class SystemValue
#Region "�v���p�e�B"

    Public Shared SysValueGet As Boolean = False
    'Public Shared SysLoginDate As Date = #12/31/2087#

    Public Shared SysWFSetPGCode As Integer = 209002

#End Region

#Region "���\�b�h"
    ''' <summary>SetSystemSettingWF -- WF�p�ϐ��l�̓W�J</summary>
    ''' <param name="dr">SystemLoad�Ŏ擾�����V�X�e���f�[�^��</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2008.07.16" name="��t �F��">N��SetSystemSetting���AWF�����̂ݕ����������\�b�h</history>
    ''' <history version="2" date="2008.07.29" name="��t �F��">WF�ݒ�̃v���O�����R�[�h��ǉ�</history>
    Public Shared Sub SetSystemSettingWF(ByVal dr As DataRow)

        'SysLoginDate = fToDate(dr("SysLoginDate"))

        'SysWFSetPGCode = fToInt(dr("SysWFSetPGCode"))

    End Sub

#End Region

End Class
