
#Region "imports�錾"
'Imports SiU.Process.Common.Functions.CommonMethod

#End Region


''' <summary>
''' �V�X�e���ϐ�
''' </summary>
''' <remarks></remarks>
''' <history version="5.0.0.0" date="2008/04/07" name="�{�{ �~��">�V�K�쐬</history> 
''' <history version="5.0.0.0" date="2008/04/08" name="��t �F��">����w���Ŏg�p���鏉���l�ݒ荀�ڂ�ǉ�</history>
''' <history version="5.0.0.0" date="2008/04/08-2" name="��t �F��">SysUserNo�ǉ�</history>
''' <history version="5.0.0.0" date="2008/04/22" name="�Ύ� ���m">SysMindate��1000/01/01�ɏC���ASysDefCodeN,SysDefCodeX��ǉ�</history>
''' <history version="5.0.0.0" date="2008/05/01" name="�Ύ� ���m">�����R�[�h���l�^�p������(SysCharKeta)�̒ǉ�</history>
''' <history version="5.0.0.0" date="2008.07.16" name="��t �F��">WF�p�����t�@�C���𐶐�</history>
''' <history version="5.0.0.0" date="2009.02.25" name="���V ��T">����R�[�h�̕����񉻑Ή�</history>
''' <history version="5.0.0.0" date="2009.03.28" name="���c �v�`">SysMinCode,SysMaxCode�̕����^SysMinCodeX,SysMaxCodeX��ǉ�</history>
''' <history version="5.0.0.0" date="2009.05.27" name="���c �v�`">SysBmnKaisou�̓��͗p(SysBmnKaisouIn)��ǉ�</history>
''' <history version="5.0.0.0" date="2009.11.05" name="�{�{ �~��">����p���ʕϐ��̒ǉ�</history>
''' <history version="5.0.0.0" date="2010.03.05" name="�{�{ �~��">����p���ʕϐ��̊�{�\�Z�ԍ��̒ǉ�</history>
''' <history version="5.0.1.1" date="2011.12.15" name="���c �v�`">�A�Z���u�����E�Q�Ɛݒ�𐳂����ݒ肵����</history>
''' <history version="5.0.2.0" date="2012.09.15" name="���c �v�`">�_�~�[�̃R���X�g���N�^��ǉ�</history>
Public Class SystemValue
#Region "�v���p�e�B"
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

    ''2009.02.25 UPD START ISHIZAWA ����R�[�h�̕����R�[�h���Ή�
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
    '' �Ǘ��}�X�^�@CODE=1 �A�C�e��
    Public Shared SysDateFormat As Integer = 0

    Public Shared SysDatePrint As Integer = 0
    Public Shared SysKaisPrint As Integer = 0
    Public Shared SysReportPrint As Integer = 0
    ''==========================================================
    '' �Ǘ��}�X�^�@CODE=2 �A�C�e��
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
    ''2009.02.25 ADD START ISHIZAWA ����R�[�h�̕����R�[�h���Ή�
    '' �Ǘ��}�X�^�@CODE=4 �A�C�e��
    Public Shared SysBmnFormatType As Integer = 0
    Public Shared SysBmnFormat As String = String.Empty
    Public Shared SysMinBmn As String = String.Empty
    Public Shared SysMaxBmn As String = "ZZZZZZZZZ"
    ''2009.02.25 ADD END

#End Region

#Region " �����ϐ�"
    Private Ret As Boolean              ' ����p
#End Region

#Region "�R���X�g���N�^"

    ''2012.09.15 ADD START
    ''' <summary>
    ''' New
    ''' </summary>
    ''' <remarks>�_�~�[</remarks>
    ''' <history version="1" date="2012.09.15" name="���c �v�`">�V�K�쐬</history>
    Public Sub New()

    End Sub
    ''2012.09.15 ADD END

    ''' <summary>
    ''' New
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <remarks></remarks>
    ''' <history version="1" date="2008/04/07" name="�{�{ �~��">�V�K�쐬</history>
    Public Sub New(ByVal dr As DataRow)

        ''�z��Ɏ擾�l��ݒ肵�܂��B
        'SetSystemSetting(dr)

    End Sub

#End Region

    '#Region "���\�b�h"
    '    ''' <summary>
    '    ''' �ϐ��l�̓W�J
    '    ''' </summary>
    '    ''' <param name="dr">SystemLoad�Ŏ擾�����V�X�e���f�[�^��</param>
    '    ''' <remarks></remarks>
    '    ''' <history version="1" date="2008/04/07" name="�{�{ �~��">�V�K�쐬</history> 
    '    ''' <history version="2" date="2008/04/08" name="��t �F��">SysUserNo��ǉ�</history>
    '    ''' <history version="3" date="2008/04/22" name="�Ύ� ���m">SysDefCodeN,SysDefCodeX��ǉ�</history> 
    '    ''' <history version="4" date="2008/05/01" name="�Ύ� ���m">�����R�[�h���l�^�p������(SysCharKeta)�̒ǉ�</history>
    '    ''' <history version="5" date="2008/06/23" name="�ޗ� �F�u">SysDDate�̎����������O</history>
    '    ''' <history version="6" date="2008.07.16" name="��t �F��">WF�ݒ�p�̌Ăяo�����\�b�h��ǉ�</history>
    '    ''' <history version="7" date="2009.01.07" name="���R �r�j">�̔��Ǘ��ݒ�p�̌Ăяo�����\�b�h��ǉ�</history>
    '    ''' <history version="8" date="2009.02.25" name="���V ��T">����R�[�h�̕����R�[�h���Ή�</history>
    '    ''' <history version="9" date="2009.03.28" name="���c �v�`">SysMinCode,SysMaxCode�̕����^SysMinCodeX,SysMaxCodeX��ǉ�</history>
    '    ''' <history version="10" date="2009.05.27" name="���c �v�`">SysBmnKaisou�̓��͗p(SysBmnKaisouIn)��ǉ�</history>
    '    ''' <history version="11" date="2009.11.05" name="�{�{ �~��">����ݒ�p�̌Ăяo�����\�b�h��ǉ�</history>
    '    Public Shared Sub SetSystemSetting(ByVal dr As DataRow)

    '        SysOpid = fToText(dr("SysOpid"))
    '        SysUserNo = fToInt(dr("SysUserNo"))
    '        SysKais = fToInt(dr("SysKais"))
    '        ''2008.06.23 UPD START T.Nara �����������O
    '        'SysDDate = fToDate(dr("SysDDate"))
    '        SysDDate = fToDate(dr("SysDDate"), #1/1/1000#, BrowsefToDate.fToDate_DateOnly)
    '        ''2008.06.23 UPD END
    '        SysLoginTime = fToDate(dr("SysLoginTime"))
    '        SysLogoutTime = fToDate(dr("SysLogoutTime"))
    '        SysMenuPId = fToInt(dr("SysMenuPId"))
    '        SysDnoStart = fToInt(dr("SysDnoStart"))
    '        SysDnoEnd = fToInt(dr("SysDnoEnd"))
    '        SysTant = fToInt(dr("SysTant"))

    '        ''2009.02.25 UPD START ISHIZAWA ����R�[�h�̕����R�[�h���Ή�
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
    '        ''2009.02.25 ADD START ISHIZAWA ����R�[�h�̕����R�[�h���Ή�
    '        SysBmnFormatType = fToInt(dr("SysBmnFormatType"))
    '        ''2009.02.25 ADD END

    '        SysCharKeta = fToInt(dr("SysCharKeta"))

    '        SysSKojFormat = fToText(dr("SysSKojFormat"))
    '        SysKojFormat = fToText(dr("SysKojFormat"))
    '        SysKojUwFormat = fToText(dr("SysKojUwFormat"))
    '        SysTorFormat = fToText(dr("SysTorFormat"))
    '        ''2009.02.25 ADD START ISHIZAWA ����R�[�h�̕����R�[�h���Ή�
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
    '        ''2009.02.25 ADD START ISHIZAWA ����R�[�h�̕����R�[�h���Ή�
    '        SysMinBmn = fToText(dr("SysMinBmn"))
    '        SysMaxBmn = fToText(dr("SysMaxBmn"))
    '        ''2009.02.25 ADD END
    '        SysKojUwKihon = fToText(dr("SysKojUwKihon"))

    '        SysDateFormat = fToInt(dr("SysDateFormat"))
    '        SysDatePrint = fToInt(dr("SysDatePrint"))
    '        SysKaisPrint = fToInt(dr("SysKaisPrint"))
    '        SysReportPrint = fToInt(dr("SysReportPrint"))

    '        ' '' WF�Ŏg�p����ݒ���Z�b�g
    '        'SetSystemSettingWF(dr)

    '        ' ''2009.01.07 ADD START NAKAYAMA
    '        ' '' �̔��Ŏg�p����ݒ���Z�b�g
    '        'SetSystemSettingHanbai(dr)
    '        ' ''2009.01.07 ADD END

    '        ' ''2009.11.05 ADD START MIYAMOTO
    '        'SetSystemSettingGenba(dr)
    '        ' ''2009.11.05 ADD END
    '    End Sub

    '#End Region

End Class
