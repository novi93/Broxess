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
''' CommonMethod -- ���ʊ֐�
''' </summary>
''' <remarks></remarks>
''' <history version="5.0.0.0" date="" name="">���ŕs��</history>
''' <history version="5.0.0.0" date="2009.03.02" name="��� �o�u��">����R�[�h�̕����R�[�h���Ή�</history>
''' <history version="5.0.0.0" date="2009.03.25" name="��t �F��">�֐��̈�����Nothing�̏ꍇ�A�N���X�̏����l�iMinValue�j�Ƃ��Ĉ����悤�ɏC��</history>
''' <history version="5.0.0.0" date="2009.06.01" name="���� ����">�؂蔲����������̃o�C�g���擾�s�����C��</history>
''' <history version="5.0.0.0" date="2009.08.13" name="���c �v�`">�N�Ȃ��̓��t�����ɂ��J���`�����w��</history>
''' <history version="5.0.0.0" date="2010.03.04" name="�Ĉ� �D��">fSqlStrType��ǉ�</history>
''' <history version="5.0.0.0" date="2010.03.30" name="���� ����">���i���萔��ǉ�</history>
''' <history version="5.0.0.0" date="2010.04.08" name="�Ĉ� �D��">ZeiCalculate��ǉ�</history>
''' <history version="5.0.0.0" date="2010.05.13" name="���� ����">���匠���Ɋւ��鋤�ʃ��\�b�h��ǉ�</history>
''' <history version="5.0.0.0" date="2010.05.26" name="���� ����">(�ۑ�\:3091)SiuFormat�̘a��̈����ŃG���[���������[�v����̂��C��</history>
''' <history version="5.0.0.0" date="2010.07.02" name="���� ����">fSqlStr��overload����DataColumn����^�EPROCESS�ł̗p�r�𔻒f���Ēu�����鏈����ǉ�
'''                                                               DataRow����INSERT/UPDATE/DELETE�𐶐�����fSqlStrInsert����ǉ�</history>
''' <history version="5.0.0.0" date="2010.07.26" name="HANH-LV">����V�X�e���I�t���C���Ή�</history>
''' <history version="5.0.1.1" date="2011.12.14" name="���c �v�`">(SC�ŗ����錴���ɂȂ��Ă�����)EventLog�ւ̏������݃G���[�͉������悤���Ȃ��̂ŃX���[����</history>
''' <history version="5.0.1.1" date="2014.03.12" name="���c �v�`">�����̏����w�肪�s���ȈׁA12���Ԑ��Ńt�H�[�}�b�g�����(�Q�Ɩ�肪�������邩������Ȃ��̂Ńo�[�W�����͏グ�܂���)</history>
''' <history version="5.0.1.1" date="2014.06.26" name="TRI-PQ">(TID.3220-IID.26713)SSL��(�Q�Ɩ�肪�������邩������Ȃ��̂Ńo�[�W�����͏グ�܂���)</history>
''' <history version="6.0.0.0" date="2017.08.24" name="HAI-NM">SQL��PROCES.S5����SERVER2016(�v���W�F�N�g����)�Ɉڍs����B</history>
Public Class CommonMethod

    ''2010.03.30 ADD START SHIMIZU
    ''' <summary>
    ''' ProductName -- ���i��
    ''' </summary>
    ''' <remarks>�ύX�֎~</remarks>
    Public Const ProductName As String = "Ussol.Process"
    ''2010.03.30 ADD END

    '' 2010.07.26 ADD START HANH-LV
    '' �Ή����e�F����V�X�e���I�t���C���Ή�
    ''' <summary>�I�t���C�����[�h�̎���</summary>
    ''' <remarks> </remarks>
    Public Shared IsOfflineMode As Boolean = False
    '' 2010.07.26 ADD END HANH-LV

#Region " �񋓑� "
    ''***************************************************************************************************************************************************

    ''' <summary>BrowsefSqlStr -- ORACLE SQL��C���p�ϐ�</summary>
    ''' <remarks></remarks>
    Public Enum BrowsefSqlStr
        ''' <summary>BigInt�^</summary>
        fSqlStr_BigInt = 16
        ''' <summary>Binary�^</summary>
        fSqlStr_Binary = 9
        ''' <summary>Boolean�^</summary>
        fSqlStr_Boolean = 1
        ''' <summary>Byte�^</summary>
        fSqlStr_Byte = 2
        ''' <summary>Char�^</summary>
        fSqlStr_Char = 18
        ''' <summary>Currency�^</summary>
        fSqlStr_Currency = 5
        ''' <summary>DATE�^</summary>
        fSqlStr_DATE = 8
        ''' <summary>Decimal�^</summary>
        fSqlStr_Decimal = 20
        ''' <summary>Double�^</summary>
        fSqlStr_Double = 7
        ''' <summary>Float�^</summary>
        fSqlStr_Float = 21
        ''' <summary>GUID�^</summary>
        fSqlStr_GUID = 15
        ''' <summary>Integer�^</summary>
        fSqlStr_Integer = 3
        ''' <summary>Long�^</summary>
        fSqlStr_Long = 4
        ''' <summary>LongBinary�^</summary>
        fSqlStr_LongBinary = 11
        ''' <summary>Memo�^</summary>
        fSqlStr_Memo = 12
        ''' <summary>Numeric�^</summary>
        fSqlStr_Numeric = 19
        ''' <summary>Single�^</summary>
        fSqlStr_Single = 6
        ''' <summary>Text�^</summary>
        fSqlStr_Text = 10
        ''' <summary>Time�^</summary>
        fSqlStr_Time = 22
        ''' <summary>TimeStamp�^</summary>
        fSqlStr_TimeStamp = 23
        ''' <summary>VarBinary�^</summary>
        fSqlStr_VarBinary = 17
        ''' <summary>DateTime�^</summary>
        fSqlStr_DateTime = 51
    End Enum

    ''' <summary>BrowsefToNumber -- fToNumber�ԋp�l�̐��l����</summary>
    ''' <remarks></remarks>
    Public Enum BrowsefToNumber
        ''' <summary>Integer�^�ɕϊ�</summary>
        fToNumber_INTEGER = 1
        ''' <summary>Long�^�ɕϊ�</summary>
        fToNumber_LONG = 2
        ''' <summary>Single�^�ɕϊ�</summary>
        fToNumber_SINGLE = 3
        ''' <summary>Double�^�ɕϊ�</summary>
        fToNumber_DOUBLE = 4
        ''' <summary>Decimal�^�ɕϊ�</summary>
        fToNumber_DECIMAL = 5
        ''' <summary>Short�^�ɕϊ�</summary>
        fToNumber_SHORT = 6
    End Enum

    ''' <summary>BrowsefToDate -- fToDate�ԋp�l�̐��l����</summary>
    ''' <remarks></remarks>
    Public Enum BrowsefToDate
        ''' <summary>���t�����ŕԋp</summary>
        fToDate_DateTime = 1
        ''' <summary>���t�݂̂ŕԋp</summary>
        fToDate_DateOnly = 2
        ''' <summary>�����݂̂ŕԋp</summary>
        fToDate_TimeOnly = 3
    End Enum

    ''' <summary>BrowseHasu -- �[�������敪</summary>
    ''' <remarks></remarks>
    Public Enum BrowseHasu
        ''' <summary>�؂�グ</summary>
        bHasu_Kiriage = 1
        ''' <summary>�؎̂�</summary>
        bHasu_Kirisute = 2
        ''' <summary>�l�̌ܓ�</summary>
        bHasu_Shisyagonyu = 3
    End Enum

    ''' <summary>
    ''' ���x����ԋ敪
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum BrouseConfig
        ''' <summary>���o�^����</summary>
        NoData = 1
        ''' <summary>�V�K�o�^����(���͊J�n�O)</summary>
        NewData = 2
        ''' <summary>�V�K�o�^����(���͊J�n��)</summary>
        NewDataInput = 3
        ''' <summary>�C���o�^����(���͊J�n�O)</summary>
        UpdData = 4
        ''' <summary>�C���o�^����(���͊J�n��)</summary>
        UpdDataInput = 5
        ''' <summary>������������</summary>
        NoSec = 6
        ''' <summary>�X�V������</summary>
        Running = 7
        ''' <summary>������(���x���̔�\��)</summary>
        Initialize = 99
    End Enum

    ''' <summary>ControlKbn -- �R���g���[���敪</summary>
    ''' <remarks>�R���g���[�����P�Ƃ܂��͔͈́i�J�n/�I���j�̔���t���O</remarks>
    ''' <history version="1" date="2008.05.23" name="��t �F��">�V�K�쐬</history>
    Public Enum ControlKbn As Integer
        ''' <summary>Only</summary>
        ''' <remarks>�P�Ǝw��</remarks>
        Only = 0
        ''' <summary>From</summary>
        ''' <remarks>�J�n�l�w��</remarks>
        From = 1
        ''' <summary>To</summary>
        ''' <remarks>�I���l�w��</remarks>
        [To] = 2
    End Enum

    ''' <summary>SysCodeType -- �V�X�e���R�[�h�敪</summary>
    ''' <remarks>�e��V�X�e���R�[�h�iSysMin***/SysMax***�j������</remarks>
    ''' <history version="1" date="2008.05.23" name="��t �F��">�V�K�쐬</history>
    ''' <history version="2" date="2009.03.02" name="��� �o�u��">����R�[�h�̕����R�[�h���Ή�</history>
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
    ''' UpdType -- �X�V�敪(1:�V�K, 2:�X�V, 3:�폜)
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="���� ����">�V�K�쐬</history>
    Public Enum UpdType As Integer
        Insert = 1
        Update = 2
        Delete = 3
    End Enum

    ''' <summary>ActionType</summary>
    ''' <remarks>�A�N�V�����萔</remarks>
    ''' <history version="1" date="2008.06.20" name="��t �F��">�V�K�쐬</history>
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

        '' Func�L�[���Ή��̱���ݒ萔��100�`
        ExportHTML = 101

    End Enum

    ''2010.05.13 ADD START SHIMIZU ���匠���Ɋւ��鋤�ʃ��\�b�h�̈�����Boolean�ł͂Ȃ�Integer(�񋓑̎g�p)�ɏC��
    ''' <summary>
    ''' CheckBmnSecType -- ���匠���Ɋւ��鋤�ʃ��\�b�h�̈���
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history version="2" date="2010.05.13" name="���� ����">�V�K�쐬</history>
    Public Enum CheckBmnSecType As Integer
        Sansyo = 1
        Input = 2
    End Enum
    ''2010.05.13 ADD END

#End Region

#Region " �u���֐� "
    ''***************************************************************************************************************************************************

    ''' <summary>fReplaceString -- ������̒��̕������u�����܂��B</summary>
    ''' <param name="pString">�������镶����</param>
    ''' <param name="pShString">�������镶����</param>
    ''' <param name="pRepString">�u�����镶����</param>
    ''' <param name="fCaseSensitive">�啶���^�����������  ����ꍇ��True�A��ʂ��Ȃ��ꍇ��False</param>
    ''' <returns>�u����̕�����</returns>
    ''' <remarks>������̒��̕������u�����܂��B</remarks>
    ''' <�쐬��>Total VB SourceBook 6</�쐬��>
    ''' <���l></���l>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
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

        '' �^����ꂽ�p�����[�^�����r�^�C�v��ݒ�
        If fCaseSensitive Then
            'StringComparison = System.StringComparison.CurrentCulture

            '�啶��/�������̋�ʂ�����ꍇ��String.Replace���g�p����
            Return pString.Replace(pShString, pRepString)
        Else
            StringComparison = System.StringComparison.CurrentCultureIgnoreCase
        End If
        strTmp = pString

        Try

            '' �ŏ��̌������s���܂��B
            intPos = strTmp.IndexOf(pShString, 0, StringComparison)

            If pRepString.Length = 0 Then
                strTmp = pString.Replace(pShString, pRepString)
            Else
                Do While intPos > -1
                    ' ������̌������J��Ԃ��܂��B
                    strTmp = strTmp.Substring(0, intPos) & pRepString & strTmp.Substring(intPos + pRepString.Length, strTmp.Length - (intPos + pRepString.Length))
                    ' ���̌������s���܂��B
                    intPos = strTmp.IndexOf(pShString, intPos + pRepString.Length, StringComparison)
                Loop
            End If


        Catch ex As Exception

            Throw New Exception("fReplaceString:������u���Ɏ��s���܂����" & vbCrLf & ex.Message)
            strTmp = pString

        End Try

        '' �l��Ԃ��܂��B
        Return strTmp

    End Function

    ''' <summary>fReplaceString -- ������̒��̕������u�����܂��B</summary>
    ''' <param name="pString">�������镶����</param>
    ''' <param name="pShString">�������镶����</param>
    ''' <param name="pRepString">�u�����镶����</param>
    ''' <param name="pOption">�啶���^�����������  ����ꍇ��0�A��ʂ��Ȃ��ꍇ��1</param>
    ''' <returns>�u����̕�����</returns>
    ''' <remarks>
    ''' ������̒��̕������u�����܂��B
    ''' ��fChangeStr�̖��̂�ύX�i�����@�\�͓������O�Ɏ����j
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Function fReplaceString(ByVal pString As String, ByVal pShString As String, ByVal pRepString As String, ByVal pOption As Integer) As String

        Return fReplaceString(pString, pShString, pRepString, (pOption = 0))

    End Function

    ''' <summary>fSqlStr -- �I���N��SQL������</summary>
    ''' <param name="pSourceString">���ϐ�</param>
    ''' <param name="pSourceStringType">�ϐ��^�C�v</param>
    ''' <returns>�I���N��SQL�ɑg�ݍ��߂镶����</returns>
    ''' <remarks>�f�[�^�^�C�v�ʂɃI���N��sql��������쐬���Ė߂��܂�</remarks>
    ''' <�쐬��>�Ĉ� �D��</�쐬��>
    ''' <���l></���l>
    ''' <history version ="2" date ="2009.06.08" name ="�ɓ� ����">�N���C�A���g�̓��t�����ɉe�������̂œ��t�ϊ��ɃJ���`�����w��</history>
    ''' <history version ="3" date ="2010.07.26" name="HANH-LV">�Ή����e�F����V�X�e���I�t���C���Ή�</history>
    ''' <history version ="4" date ="2017.08.24" name="HAI-NM">SQL��PROCES.S5����SERVER2016(�v���W�F�N�g����)�Ɉڍs����B</history>
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
    ''' <param name="pKeyCode">�ϊ��Ώەϐ�</param>
    ''' <returns>�ϊ��㕶����</returns>
    ''' <remarks></remarks>
    ''' <history version ="1" date ="2010.04.08" name ="�Ĉ� �D��">�V�O�l�[�`���̒ǉ��i������fsqlstrType�̌Ăяo���j</history>
    Public Shared Function fSqlStr(ByVal pKeyCode As Object) As String
        Return fSqlStrType(pKeyCode)
    End Function

    ''' <summary>
    ''' fSqlStrType
    ''' </summary>
    ''' <param name="pKeyCode">�ҏW�Ώە���</param>
    ''' <returns></returns>
    ''' <remarks>�w��^�ɂ��AfSqlStr��ݒ�㕶�����ԋp</remarks>
    ''' <history version ="1" date ="2010.03.04" name="�Ĉ� �D��">�V�K�쐬</history>
    <Browsable(True), Description("�����I�Ɍ^�w�肳�ꂽ���ڂ�fSqlStr�������ĕԋp")>
    Public Shared Function fSqlStrType(ByVal pKeyCode As Object) As String
        Dim xTxt As String
        If TypeOf pKeyCode Is String Then       ''������
            xTxt = fSqlStr(pKeyCode, BrowsefSqlStr.fSqlStr_Text)
        ElseIf TypeOf pKeyCode Is Date Then     ''���t
            xTxt = fSqlStr(pKeyCode, BrowsefSqlStr.fSqlStr_DATE)
        Else                                    ''����ȊO�͐��l
            xTxt = fSqlStr(pKeyCode, BrowsefSqlStr.fSqlStr_Numeric)
        End If
        Return xTxt

    End Function

    ''' <summary>fDtSqlStr</summary>
    ''' <param name="pSourceString">���ϐ�</param>
    ''' <param name="pSourceStringType">�ϐ��^�C�v</param>
    ''' <returns></returns>
    ''' <remarks>DataTable Sql������</remarks>
    ''' <history version ="1" date ="2007.03.05" name ="�Ĉ� �D��">V4fJetSqlStr���ڍs</history>
    ''' <history version ="2" date ="2008.10.01" name ="�Ĉ� �D��">SQL������̋K��MDB��DataTable�ňႤ���߂ɏ��������C</history>
    ''' <history version ="3" date ="2009.03.25" name ="��t �F��">���������Nothing�����ǉ�</history>
    ''' <history version ="4" date ="2009.06.08" name ="�ɓ� ����">�N���C�A���g�̓��t�����ɉe�������̂œ��t�ϊ��ɃJ���`�����w��</history>
    ''' <history version ="5" date ="2014.03.12" name ="���c �v�`">�����̏����w�肪�s���ȈׁA12���Ԑ��Ńt�H�[�}�b�g�����</history>
    <DebuggerStepThrough(), Description("DataTable Sql������")>
    Public Shared Function fDtSqlStr(ByVal pSourceString As Object, ByVal pSourceStringType As BrowsefSqlStr) As String
        Dim xRet As String

        If IsDBNull(pSourceString) Or pSourceString Is Nothing Then
            xRet = "NULL"
        Else
            Select Case pSourceStringType
                Case BrowsefSqlStr.fSqlStr_DateTime
                    ''2009.06.08 UPD START ITO �N���C�A���g�̓��t�����ɉe�������̂ŃJ���`�����w��
                    'xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy hh:mm:ss") & "#"
                    ''2014.03.12 UPD START
                    'xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy hh:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) & "#"
                    xRet = "#" & CDate(pSourceString).ToString("MM/dd/yyyy HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) & "#"
                    ''2014.03.12 UPD END
                    ''2009.06.08 UPD END ITO
                Case BrowsefSqlStr.fSqlStr_DATE
                    ''2009.06.08 UPD START ITO �N���C�A���g�̓��t�����ɉe�������̂ŃJ���`�����w��
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

    ''' <summary>fChangeTilde -- ������̒���"�`"�̃R�[�h��ϊ�����</summary>
    ''' <param name="pString">�������镶����</param>
    ''' <returns>�u����̕�����</returns>
    ''' <remarks>��H301C �� ��HFF5E �ɕϊ�����</remarks>
    ''' <�쐬��>�Ĉ� �D��</�쐬��>
    ''' <���l>Oracle vs Microsoft ��݊�</���l>
    <DebuggerStepThrough()>
    Public Shared Function fChangeTilde(ByVal pString As String) As String
        pString = pString.Replace(ChrW(&H301C), ChrW(&HFF5E))
        Return pString
    End Function

    ''' <summary>fCodeformat -- ���l���w�茅���̉E�l�ߕ����ł��ǂ��</summary>
    ''' <param name="pCode">���l</param>
    ''' <param name="pKeta">�w�茅��</param>
    ''' <param name="pFormat">�t�H�[�}�b�g����(Format�֐�����)</param>
    ''' <returns>�E�l�ߕ���</returns>
    ''' <remarks>�w�萔�l���w�茅���̕������ŕ�����Ƃ��ĕԂ�</remarks>
    ''' <�쐬��>����@�b�j</�쐬��>
    ''' <���l></���l>
    <DebuggerStepThrough()>
    Public Shared Function fCodeformat(ByVal pCode As Object, Optional ByVal pKeta As Integer = 9, Optional ByVal pFormat As String = "") As String
        Dim nP As Decimal = fToNumberEx(pCode)
        Dim xS As String
        Dim ResultStr As String

        ''�����w�肪����ꍇ�͏�����ݒ肷��
        If pFormat = "" Then
            xS = Format$(nP)
        Else
            xS = Format$(nP, pFormat)
        End If

        ''�w�茅���ɂ܂Ƃ߂�
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

    ''' <summary>SiuFormat -- �a��Ή��t�H�[�}�b�g�֐�</summary>
    ''' <param name="Expression">�ϊ��ΏۃI�u�W�F�N�g</param>
    ''' <param name="Style">�ϊ������t�H�[�}�b�g</param>
    ''' <returns>�ϊ��㕶����</returns>
    ''' <remarks></remarks>
    ''' <history version ="2" date ="2009.06.08" name ="�ɓ� ����">�N���C�A���g�̓��t�����ɉe�������̂œ��t�ϊ��ɃJ���`�����w��</history>
    ''' <history version ="3" date ="2009.08.13" name ="���c �v�`">�N�Ȃ��̓��t�����ɂ��J���`�����w��</history>
    ''' <history version ="4" date ="2009.08.19" name ="��t �F��">����w�莞�ɃJ�����g�������g�p����悤�ɕύX</history>
    ''' <history version ="5" date ="2009.08.20" name ="��t �F��">ver4�𖳌��ɂ��āA�����C�N</history>
    ''' <history version ="6" date ="2010.05.26" name ="���� ����">(�ۑ�\:3091)SiuFormat�̘a��̈����ŃG���[���������[�v����̂��C��</history>
    <DebuggerStepThrough()>
    Public Shared Function SiuFormat(ByVal Expression As Object, Optional ByVal Style As String = "") As String
        Dim cul As New CultureInfo("ja-JP")
        cul.DateTimeFormat.Calendar = New JapaneseCalendar
        Dim ExpressionDate As Date
        Dim strReturn As String = ""
        Dim strEra As String = ""

        Dim eraKanji() As String = {"", "��", "��", "��", "��"}
        Dim eraAlpha() As String = {"", "M", "T", "S", "H"}

        If IsDate(Expression) Then
            Expression = CType(Expression, Date)
        End If

        If TypeOf (Expression) Is Date Then ''���t�̏ꍇ��
            Try
                ExpressionDate = DirectCast(Expression, Date)
                '' 2009.08.20 UPD START
                '    If Style Like "*gggee*" Then                            '����00'
                '        Style = Style.Replace("gggee", "ggyy")
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '    ElseIf Style Like "*ggge*" Then                         '����0'
                '        Style = Style.Replace("ggge", "ggy")
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '    ElseIf Style Like "*ggee*" Then                         '��00'
                '        Style = Style.Replace("ggee", "yy")
                '        ''�����擾
                '        strEra = eraKanji(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    ElseIf Style Like "*gge*" Then                          '��0'
                '        Style = Style.Replace("gge", "y")
                '        ''�����擾
                '        strEra = eraKanji(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    ElseIf Style Like "*gee*" Then                          'H00'
                '        Style = Style.Replace("gee", "yy")
                '        ''�����擾
                '        strEra = eraAlpha(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    ElseIf Style Like "*ge*" Then                           'H0'
                '        Style = Style.Replace("ge", "y")
                '        ''�����擾
                '        strEra = eraAlpha(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                '        strReturn = ExpressionDate.ToString(Style, cul)
                '        strReturn = strEra & strReturn
                '    Else
                '        ''2009.06.08 UPD START ITO �N���C�A���g�̓��t�����ɉe�������̂ŃJ���`�����w��
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
                ' �j�������̓J���`���Ɋւ�炸�a��Ƃ���̂Ő�ɒu��
                Style = ReplaceJapaneseDayOfWeek(ExpressionDate, Style)
                ''2010.05.26 ADD END

                '' �܂��ug�v�A�ue�v���܂܂�邩�ۂ��ŃJ���`���w�肷�邩�ۂ��𔻒f����
                If Style.ToLower.IndexOf("g") > -1 OrElse Style.ToLower.IndexOf("e") > -1 Then
                    '' �J���`���w��

                    ''�ue�v�ˁuy�v
                    Style = Style.Replace("e"c, "y"c)

                    '' ���������ɂāA�ug�v�̐����擾
                    Dim eraCount As Integer = 0
                    For Each c As Char In Style.ToCharArray
                        If c.Equals("g"c) Then
                            eraCount += 1
                        End If
                    Next

                    '' �W���őΉ��ł��Ȃ������́A�z�񂩂當���擾
                    Select Case eraCount
                        Case 1
                            ''�ug�v����
                            Style = Style.Replace("g"c, "")
                            ''�����A���t�@�x�b�g�擾
                            strEra = eraAlpha(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                        Case 2
                            ''�ug�v����
                            Style = Style.Replace("g"c, "")
                            ''�����������擾
                            strEra = eraKanji(cul.DateTimeFormat.Calendar.GetEra(ExpressionDate))
                        Case Else
                            strEra = String.Empty
                    End Select

                    '' ���� + �W���Ŗ߂��i���������񂪂Ȃ��ꍇ�͕W�������őΉ��j
                    strReturn = strEra & ExpressionDate.ToString(Style, cul)

                Else
                    ''2010.05.26 DEL START SHIMIZU �a��ł̗�O���ɖ������[�v�̌����ɂȂ��Ă����̂ō폜
                    ' '' ���j���͂��Ȃ炸���{�����ŕԂ�
                    'Dim ddd As String = String.Empty
                    ' '' �j���w�菑���������A���炩���߃J���`���w��Ŏ擾�����������Ă���
                    'If Style.ToLower.IndexOf("dddd") > -1 Then
                    '    ddd = ExpressionDate.ToString("dddd", cul)
                    '    Style = Style.Replace("dddd", ddd)
                    'ElseIf Style.ToLower.IndexOf("ddd") > -1 Then
                    '    ddd = ExpressionDate.ToString("ddd", cul)
                    '    Style = Style.Replace("ddd", ddd)
                    'End If
                    ''2010.05.26 DEL END

                    '' �J���`�����w��
                    strReturn = ExpressionDate.ToString(Style, System.Globalization.DateTimeFormatInfo.InvariantInfo)
                End If
                '' 2009.08.20 UPD END

            Catch ex As Exception   ''���s�����炻�̂܂�
                '' 2009.05.19 ADD START
                ' ���s���͌����擾�ł��Ȃ��Ɣ��f���A����Ŗ߂�
                ' �N��������U�폜���A����4���w��
                Style = Style.Replace("g"c, "")
                Style = Style.Replace("e"c, "")
                Style = Style.Replace("y"c, "")
                Style = "yyyy" & Style
                '' 2009.05.19 ADD END
                ''2009.06.08 UPD START ITO �N���C�A���g�̓��t�����ɉe�������̂ŃJ���`�����w��
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
                '' ���������������āA�Ď��s
                strReturn = SiuFormat(Expression, Style)
            End Try
        Else
            strReturn = Format(Expression, Style)
        End If

        Return strReturn
    End Function

    ''' <summary>
    ''' ReplaceJapaneseDayOfWeek -- �j��������a��Œu��
    ''' </summary>
    ''' <param name="targetDate">�Ώۓ��t</param>
    ''' <param name="style">����</param>
    ''' <returns>�u����������</returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.05.26" name="���� ����">�V�K�쐬</history>
    <DebuggerStepThrough()>
    Public Shared Function ReplaceJapaneseDayOfWeek(ByVal targetDate As Date, ByVal style As String) As String
        Dim xDay As String  ' �j���ϐ�

        ' �j�����擾���ĕϐ��ɓ��{���ݒ�
        Select Case targetDate.DayOfWeek
            Case DayOfWeek.Sunday
                xDay = "��"
            Case DayOfWeek.Monday
                xDay = "��"
            Case DayOfWeek.Tuesday
                xDay = "��"
            Case DayOfWeek.Wednesday
                xDay = "��"
            Case DayOfWeek.Thursday
                xDay = "��"
            Case DayOfWeek.Friday
                xDay = "��"
            Case DayOfWeek.Saturday
                xDay = "�y"
            Case Else
                xDay = String.Empty
        End Select

        ' �����ɍ��킹�Ēu��
        If style.IndexOf("dddd") > -1 Then
            style = style.Replace("dddd", xDay & "�j��")
        ElseIf style.IndexOf("ddd") > -1 Then
            style = style.Replace("ddd", xDay)
        End If

        Return style
    End Function

    ''' <summary>ConvCrLf -- vbCrLf����\n</summary>
    ''' <param name="str">�ϊ��Ώە�����</param>
    ''' <param name="mode">0:vbCrLf��\n 1:\n��vbCrLf</param>
    ''' <returns>�ϊ��㕶����</returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2008.05.12" name="��t �F��">�V�K�쐬</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
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

        '' ���K�\���Ō���
        Dim regex As New System.Text.RegularExpressions.Regex(xPattern)

        '' �p�^�[���}�b�`
        If Not regex.IsMatch(str) Then
            Return str
        End If

        '' �u��
        Return regex.Replace(str, xReplace)

    End Function

    ''' <summary>
    ''' NumDateFormat
    ''' </summary>
    ''' <param name="strTextChk"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    Public Shared Function NumDateFormat(ByVal strTextChk As String) As String
        If strTextChk Is Nothing Then
            strTextChk = String.Empty
        End If
        If Len(strTextChk) = 6 Then
            '����Q���{��(2��)�{��(2��)�Ɣ��f����B
            strTextChk = CDate(String.Format("{0:0000/00/00}", CDbl("20" & strTextChk)))
        ElseIf Len(strTextChk) = 8 Then
            '����S���{��(2��)�{��(2��)�Ɣ��f����B
            strTextChk = CDate(String.Format("{0:0000/00/00}", CDbl(strTextChk)))
        End If

        Return strTextChk

    End Function

    ''' <summary>
    ''' fSqlStr -- DataColumn����^��PROCESS�ł̗p�r�𔻒f����Oracle�`����SQL�l�ɕϊ�
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="col"></param>
    ''' <param name="updSyubetu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="���� ����">�V�K�쐬</history>
    ''' <history version="2" date="2017.08.24" name="HAI-NM">SQL��PROCES.S5����SERVER2016(�v���W�F�N�g����)�Ɉڍs����B</history>
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
                    ' �^���f
                    xSqlValue = fSqlStrType(value)
                End If
        End Select

        Return xSqlValue
    End Function

#End Region

#Region " SQL�����֐� "

    ''' <summary>
    ''' fSqlStrInsert -- DataRow����Insert���𐶐�
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="���� ����">�V�K�쐬</history>
    Public Shared Function fSqlStrInsert(ByVal tableName As String, ByVal source As DataRow) As String
        ' �X�V���INSERT��SQL����
        Return fSqlStrInsert(tableName, source, UpdType.Insert)
    End Function

    ''' <summary>
    ''' fSqlStrInsert -- DataRow����Insert���𐶐�
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <param name="updSyubetu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="���� ����">�V�K�쐬</history>
    Public Shared Function fSqlStrInsert(ByVal tableName As String, ByVal source As DataRow, ByVal updSyubetu As UpdType) As String
        Dim sbSqlTxt As New System.Text.StringBuilder    ' SQL
        Dim sbFieldsTxt As New System.Text.StringBuilder ' �t�B�[���h
        Dim sbValuesTxt As New System.Text.StringBuilder ' �l

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
    ''' fSqlStrUpdate -- DataRow��Key�z�񂩂�Update���𐶐�
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <param name="keyColNames"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="���� ����">�V�K�쐬</history>
    Public Shared Function fSqlStrUpdate(ByVal tableName As String, ByVal source As DataRow, ByVal keyColNames() As String) As String
        Dim sbSqlTxt As New System.Text.StringBuilder    ' SQL
        Dim sbValuesTxt As New System.Text.StringBuilder ' �l
        Dim sbWhereTxt As New System.Text.StringBuilder  ' ����

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
    ''' fSqlStrDelete -- DataRow��Key�z�񂩂�Delete���𐶐�
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="source"></param>
    ''' <param name="keyColNames"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2010.07.02" name="���� ����">�V�K�쐬</history>
    Public Shared Function fSqlStrDelete(ByVal tableName As String, ByVal source As DataRow, ByVal keyColNames() As String) As String
        Dim sbSqlTxt As New System.Text.StringBuilder    ' SQL
        Dim sbWhereTxt As New System.Text.StringBuilder  ' ����

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

#Region " ���o�E��r�֐� "
    ''***************************************************************************************************************************************************

    ''' <summary>fLeftByte -- ���������񒊏o</summary>
    ''' <param name="pSource">�Ώە�����</param>
    ''' <param name="pLeftLen">���o�o�C�g��</param>
    ''' <param name="pCharSet">�G���R�[�h�w��: 1=shift-jis 2=Unicode(16) 3=UTF8 4=UTF7 5=UTF32</param>
    ''' <returns>���o������</returns>
    ''' <remarks>�w�蕶����̍�����w��o�C�g���𒊏o����</remarks>
    ''' <�쐬��>Process4����̕����쐬</�쐬��>
    ''' <���l>�S�p�����̔����ŏI������ꍇ�͐؂�̂Ă�</���l>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    <DebuggerStepThrough()>
    Public Shared Function fLeftByte(ByVal pSource As Object, ByVal pLeftLen As Integer, ByVal pCharSet As Integer) As String
        Dim xBuf As String = ""
        Dim xBuf2 As String = ""
        Dim i As Integer

        Try
            ''Null�̏ꍇ��Null�ŕԐM
            If pSource Is DBNull.Value Then
                Return DBNull.Value.ToString
            End If

            If pSource Is Nothing Then
                Return String.Empty
            End If

            ''�����[���̏ꍇ�͋󕶎��ԐM
            If CType(pSource, String).Length = 0 Then
                Return ""
            End If

            ''������`�F�b�N
            For i = 1 To Len(pSource) Step 1
                xBuf2 = xBuf & Mid$(CType(pSource, String), i, 1)

                'SJIS�ł̃o�C�g���Ƃ̔�r
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

    ''' <summary>fStrCompSJIS -- �������SJIS�Ŕ�r����</summary>
    ''' <param name="pxFrom">�J�n������</param>
    ''' <param name="pxTo">�I��������</param>
    ''' <returns>
    ''' True  :�J�n������ ��= �I��������
    ''' False :�J�n������ �� �I��������
    ''' </returns>
    ''' <remarks>�������SJIS�Ŕ�r����</remarks>
    ''' <�쐬��>Process4����̕����쐬</�쐬��>
    ''' <���l></���l>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
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

        'S-JIS�ւ̃G���R�[�h�I�u�W�F�N�g
        Dim SJIS_Encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        '' �ϐ��̏�����
        fStrCompSJIS = False
        nFrom = SJIS_Encode.GetBytes(pxFrom)
        'nFrom = StrConv(pxFrom, vbFromUnicode)
        nTo = SJIS_Encode.GetBytes(pxTo)
        'nTo = StrConv(pxTo, vbFromUnicode)

        '' ��r�̊J�n�|�C���^
        If LBound(nFrom) < LBound(nTo) Then
            nStart = LBound(nTo)
        Else
            nStart = LBound(nFrom)
        End If

        '' ��r�̏I���|�C���^�E�S�Ĉ�v�̌��ʒl
        If UBound(nFrom) > UBound(nTo) Then
            nEnd = UBound(nTo)
            bEQ = False
        Else
            nEnd = UBound(nFrom)
            bEQ = True
        End If

        '' ��r���[�v
        For i = nStart To nEnd
            If nFrom(i) > nTo(i) Then
                fStrCompSJIS = False
                Exit Function
            ElseIf nFrom(i) < nTo(i) Then
                fStrCompSJIS = True
                Exit Function
            End If
        Next

        '' �S�Ĉ�v�̌��ʒl
        fStrCompSJIS = bEQ

    End Function

    ''' <summary>SplitString -- ��������ꎞ���������z��ɗ��Ƃ�����</summary>
    ''' <param name="str">�Ώە�����</param>
    ''' <returns></returns>
    ''' <remarks>Split�̃Z�p���[�^�[��������</remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
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

    ''' <summary>LenB -- �o�C�g���擾</summary>
    ''' <param name="Expression">�J�E���g�Ώە�����</param>
    ''' <returns>�J�E���g����</returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Shared Function LenB(ByVal Expression As String) As Integer
        Return LenB(Expression, 1)
    End Function

    ''' <summary>LenB -- �o�C�g���擾</summary>
    ''' <param name="Expression">�J�E���g�Ώە�����</param>
    ''' <param name="pCodingNumber">�G���R�[�h�w��: 1=shift-jis 2=Unicode(16) 3=UTF8 4=UTF7 5=UTF32</param>
    ''' <returns>�J�E���g����</returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
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

    ''' <summary>SysCodeEquals -- �e��V�X�e���R�[�h�̍ŏ�/�ő��r</summary>
    ''' <param name="value">��r�Ώےl</param>
    ''' <returns>���l�̏ꍇ��True</returns>
    ''' <remarks>SysMinCode�Ƃ̔�r</remarks>
    ''' <history version="1" date="2008.05.23" name="��t �F��">�V�K�쐬</history>
    Public Shared Function SysCodeEquals(ByVal value As Object) As Boolean
        Return SysCodeEquals(value, ControlKbn.Only, SysCodeType.Code)
    End Function

    ''' <summary>SysCodeEquals -- �e��V�X�e���R�[�h�̍ŏ�/�ő��r</summary>
    ''' <param name="value">��r�Ώےl</param>
    ''' <param name="conkbn">ControlKbn�񋓑�</param>
    ''' <returns>���l�̏ꍇ��True</returns>
    ''' <remarks>SysMinCode/SysMaxCode�Ƃ̔�r</remarks>
    ''' <history version="1" date="2008.05.23" name="��t �F��">�V�K�쐬</history>
    Public Shared Function SysCodeEquals(ByVal value As Object, ByVal conkbn As Integer) As Boolean
        Return SysCodeEquals(value, conkbn, SysCodeType.Code)
    End Function

    ''' <summary>SysCodeEquals -- �e��V�X�e���R�[�h�̍ŏ�/�ő��r</summary>
    ''' <param name="value">��r�Ώےl</param>
    ''' <param name="conkbn">ControlKbn�񋓑�</param>
    ''' <param name="type">SysCodeType�񋓑�</param>
    ''' <returns>���l�̏ꍇ��True</returns>
    ''' <remarks>type�Ŏw�肳�ꂽSysMin***/SysMax***�Ƃ̔�r</remarks>
    ''' <history version="1" date="2008.05.23" name="��t �F��">�V�K�쐬</history>
    ''' <history version="2" date="2009.03.02" name="��� �o�u��">����R�[�h�̕����R�[�h���Ή�</history>
    Public Shared Function SysCodeEquals(ByVal value As Object, ByVal conkbn As Integer, ByVal type As SysCodeType) As Boolean

        Dim MinValue As Object
        Dim MaxValue As Object
        Dim CompValue As Object

        '' �^�C�v�ʂɍŏ�/�ő���擾
        Select Case type
            Case SysCodeType.Code           '' �R�[�h
                MinValue = SysMinCode
                MaxValue = SysMaxCode
            Case SysCodeType.Date           '' ���t
                MinValue = SysMinDate
                MaxValue = SysMaxDate
            Case SysCodeType.SKoj           '' �W�v�H���R�[�h
                MinValue = SysMinSKoj
                MaxValue = SysMaxSKoj
            Case SysCodeType.Koj            '' �H���R�[�h
                MinValue = SysMinKoj
                MaxValue = SysMaxKoj
            Case SysCodeType.KojUw          '' �H���ڍ׃R�[�h
                MinValue = SysMinKojUw
                MaxValue = SysMaxKojUw
            Case SysCodeType.Tor            '' �����R�[�h
                MinValue = SysMinTor
                MaxValue = SysMaxTor
            Case SysCodeType.Soko           '' �q�ɃR�[�h
                MinValue = SysMinSoko
                MaxValue = SysMaxSoko
            Case SysCodeType.Syohin         '' ���i�R�[�h
                MinValue = SysMinSyohin
                MaxValue = SysMaxSyohin
            Case SysCodeType.Jyu            '' �󒍃R�[�h
                MinValue = SysMinJyu
                MaxValue = SysMaxJyu
            Case SysCodeType.Hac            '' �����R�[�h
                MinValue = SysMinHac
                MaxValue = SysMaxHac
            Case SysCodeType.Syu            '' �o�׃R�[�h
                MinValue = SysMinSyu
                MaxValue = SysMaxSyu
            Case SysCodeType.Uri            '' ����R�[�h
                MinValue = SysMinUri
                MaxValue = SysMaxUri
            Case SysCodeType.Nyu            '' ���׃R�[�h
                MinValue = SysMinNyu
                MaxValue = SysMaxNyu
            Case SysCodeType.Sir            '' �d���R�[�h
                MinValue = SysMinSir
                MaxValue = SysMaxSir
                ''2009.03.02 ADD START OOUCHI ����R�[�h�̕����R�[�h���Ή�
            Case SysCodeType.Bmn
                MinValue = SysMinBmn
                MaxValue = SysMaxBmn
                ''2009.03.02 ADD END
            Case Else                       '' ���̑��̓R�[�h�ŏ���
                MinValue = SysMinCode
                MaxValue = SysMaxCode
        End Select

        '' ��r����敪�ɂ��A��r�l���擾
        If (conkbn = ControlKbn.Only) Or (conkbn = ControlKbn.From) Then
            '' �P�Ƃ܂��͔͈͊J�n�Ƃ̊m�F��MinValue�Ŋm�F
            CompValue = MinValue
        Else
            '' �ȊO��MaxValue�Ŋm�F
            CompValue = MaxValue
        End If

        '' ��r
        Return value.Equals(CompValue)

    End Function

#End Region

#Region " �Z�o�֐� "
    ''***************************************************************************************************************************************************

    ''' <summary>
    ''' fStartDate--���x�J�n���Z�o
    ''' </summary>
    ''' <param name="pDdate">���x</param>
    ''' <param name="pSIME">����</param>
    ''' <returns>���x�J�n��</returns>
    ''' <remarks>�w�茎�x�̊J�n�������x�{�����ŎZ�o����B</remarks>
    ''' <�쐬��>Process4����̕����쐬</�쐬��>
    ''' <���l></���l>
    <DebuggerStepThrough()>
    Public Shared Function fStartDate(ByVal pDdate As Object, ByVal pSIME As Integer) As Date
        Dim dGaitouDate As Object
        Dim dGaitouWk As Object

        ''�����lNull
        fStartDate = SysMinDate

        ''���t�ȊO�������ꍇ�͏I��
        If IsDate(pDdate) = False Then
            fStartDate = SysMinDate
            Exit Function
        End If

        ''�Y�����̓����P�ɋ����Œ�(���x�ɂ���)
        dGaitouDate = DateSerial(Year(pDdate), Month(pDdate), 1)

        ''�ߓ����R�P�̏ꍇ�͂P�����Y���Ȃ̂ł��̂܂ܖ߂��
        If pSIME = 31 Then
            fStartDate = dGaitouDate
            Exit Function
        End If

        ''�ߓ����R�P�̏ꍇ�͂P�����Y���Ȃ̂ł��̂܂ܖ߂��
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
    ''' fEndDate--���x�����Z�o
    ''' </summary>
    ''' <param name="pDdate">���x</param>
    ''' <param name="pSIME">����</param>
    ''' <returns>���x����</returns>
    ''' <remarks>�w�茎�x�̍ŏI�������x�{�����ŎZ�o����B</remarks>
    ''' <�쐬��>Process4����̕����쐬</�쐬��>
    ''' <���l></���l>
    <DebuggerStepThrough()>
    Public Shared Function fEndDate(ByVal pDdate As Object, ByVal pSIME As Integer) As Date
        Dim dGaitouDate As Object
        Dim dGaitoudd As Integer

        ''�����l=SysMinDate
        fEndDate = SysMinDate

        ''���t�ȊO�������ꍇ�͏I��
        If IsDate(pDdate) = False Then
            Exit Function
        End If

        ''�Y�����̓����P�ɋ����Œ�(���x�ɂ���)
        dGaitouDate = DateSerial(Year(pDdate), Month(pDdate), 1)

        ''�������ŏI�����Z�o
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

        ''�ŏI����31�̏ꍇ�͌��ʍŏI�����Z�o�C�ȊO�͂��̂܂ܑ��
        If dGaitoudd = 31 Then
            fEndDate = DateSerial(Year(dGaitouDate), Month(dGaitouDate) + 1, 1 - 1)
        Else
            fEndDate = DateSerial(Year(dGaitouDate), Month(dGaitouDate), dGaitoudd)
        End If
    End Function

    ''' <summary>
    ''' fProcessMonth--���x���Z�o����
    ''' </summary>
    ''' <param name="pDdate">�Y���N����</param>
    ''' <param name="pSIME">����</param>
    ''' <returns>�Ώی��x</returns>
    ''' <remarks>�Y�����t���ǂ̌��x�ɑ����邩�߂�</remarks>
    ''' <�쐬��>����@�b�j</�쐬��>
    ''' <���l></���l>
    <DebuggerStepThrough()>
    Public Shared Function fProcessMonth(ByVal pDdate As Object, ByVal pSIME As Integer) As Date
        Dim dGaitouDate As Object
        Dim nGaitoudd As Integer
        Dim nGaitouSime As Integer

        ''�����lNull
        fProcessMonth = SysMinDate

        ''���t�ȊO�������ꍇ�͏I��
        If IsDate(pDdate) = False Then
            fProcessMonth = SysMinDate
            Exit Function
        End If

        ''���x�ɂ���
        dGaitouDate = DateSerial(Year(pDdate), Month(pDdate), 1)
        nGaitoudd = Microsoft.VisualBasic.Day(pDdate)

        ''������蔻�f
        If pSIME = 31 Then
            nGaitouSime = 31
        ElseIf (pSIME >= 1) And (pSIME < 31) Then
            nGaitouSime = pSIME
        Else
            fProcessMonth = SysMinDate
            Exit Function
        End If

        ''���������������f
        If nGaitoudd <= nGaitouSime Then
            fProcessMonth = dGaitouDate
        Else
            fProcessMonth = DateSerial(Year(dGaitouDate), Month(dGaitouDate) + 1, 1)
        End If
    End Function

    ''' <summary>fKishuMonth -- ����N�����o��</summary>
    ''' <param name="pNengetu">��ʂŎw�肳�ꂽ�N��</param>
    ''' <param name="pStart_YM">���ƊJ�n�N��</param>
    ''' <returns>����N��</returns>
    ''' <remarks>��ʂ̔N���Ǝ��ƊJ�n�N���������̌������߂�</remarks>
    ''' <history version="" date="" name="��t �F��">Process4����̕���</history>
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

    ''' <summary>fRound -- �[������</summary>
    ''' <param name="pKin">�Ώۋ��z</param>
    ''' <param name="pHasuu">�[�������敪</param>
    ''' <param name="pTani">�ۂߒP��(1�ŏ����_�ȉ��C10��10�~�ȉ�)</param>
    ''' <returns>�v�Z����</returns>
    ''' <remarks>�w��P�ʂŐ��l�̊ۂߏ������s���܂��B</remarks>
    ''' <�쐬��>Process4����̕����쐬</�쐬��>
    ''' <���l></���l>
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
            Case BrowseHasu.bHasu_Kiriage       '�؂�グ
                nChukan2 = -System.Math.Sign(nChukan1) * Int(-System.Math.Abs(nChukan1))
            Case BrowseHasu.bHasu_Kirisute      '�؂�̂�
                nChukan2 = Fix(nChukan1)
            Case BrowseHasu.bHasu_Shisyagonyu   '�l�̌ܓ�
                nChukan2 = System.Math.Sign(nChukan1) * Int(System.Math.Abs(nChukan1) + 0.5D)
            Case Else
                nChukan2 = 0
        End Select

        Return nChukan2 * nTani

    End Function

    ''' <summary>
    ''' ZeiCalculate
    ''' </summary>
    ''' <param name="pKin">�Ώۋ��z</param>
    ''' <param name="pZeiInput">�œ��͋敪</param>
    ''' <param name="pZeiRitsu">�ŗ�</param>
    ''' <param name="pHasu">�[�������敪</param>
    ''' <returns>�Ŋz</returns>
    ''' <remarks></remarks>
    ''' <history version ="1" date ="2010.04.08" name="�Ĉ� �D��">�V�K�쐬</history>
    <DebuggerStepThrough(), Browsable(True), Description("�ŎZ�o pKin:�Ώۋ��z pZeiInput:1=�O��2=���� pZeiRitsu:�ŗ� pHasu:�[�������敪")>
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

#Region " �ϊ��֐� "
    ''***************************************************************************************************************************************************

    ''' <summary>fNvl -- Null�u������</summary>
    ''' <param name="pValue">�Ώەϐ�</param>
    ''' <param name="pRepValue">Null���u�������ϐ�</param>
    ''' <returns>Null�u��������̒l</returns>
    ''' <remarks>pValue��Nothing�܂���Null�̏ꍇ��pRepValue���C����ȊO��pValue�����̂܂ܕԂ�</remarks>
    <DebuggerStepThrough()>
    Public Shared Function fNvl(ByVal pValue As Object, ByVal pRepValue As Object) As Object
        If pValue Is Nothing OrElse pValue Is DBNull.Value Then
            Return pRepValue
        Else
            Return pValue
        End If
    End Function

    ''' <summary>fNvlEx -- Null�u������</summary>
    ''' <param name="pValue">�Ώەϐ�</param>
    ''' <param name="pRepValue">Null���u�������ϐ�</param>
    ''' <remarks>
    ''' pValue��Nothing�܂���Null�̏ꍇ��pRepValue���C����ȊO��pValue�����̂܂ܕԂ�
    ''' Option Strict = True �̑Ή�
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Sub fNvlEx(ByRef pValue As Object, ByVal pRepValue As Object)
        If pValue Is Nothing OrElse pValue Is DBNull.Value Then
            pValue = pRepValue
        End If
    End Sub

    ''' <summary>fToNumber -- ���l��</summary>
    ''' <param name="pVal">�Ώەϐ�</param>
    ''' <param name="pRetType">�ԋp�l�̐��l���� �ȗ�����Decimal�ɂĕԋp</param>
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
        Err.Raise(vbObjectError + 100, , "fToNumber:���l�ϊ��Ɏ��s���܂����" & vbCrLf & xErrm)
    End Function

    ''' <summary>fToNumberEx -- ���l���@�@���^�w��ȗ�</summary>
    ''' <param name="pVal">�Ώەϐ�</param>
    ''' <returns>Decimal�^�̐��l</returns>
    ''' <remarks>
    ''' Option Strict = True �̑Ή�
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Function fToNumberEx(ByVal pVal As Object) As Decimal

        '' Decimal�^���w�肵�ď�������
        Return fToNumberEx(Of Decimal)(pVal)

    End Function

    ''' <summary>fToNumberEx -- ���l��</summary>
    ''' <typeparam name="NumericType">�ԋp�l�̐��l����</typeparam>
    ''' <param name="pVal">�Ώەϐ�</param>
    ''' <returns>NumericType�Ŏw�肳�ꂽ�^�̐��l</returns>
    ''' <remarks>
    ''' Option Strict = True �̑Ή�
    ''' </remarks>
    <DebuggerStepThrough()>
    Public Shared Function fToNumberEx(Of NumericType)(ByVal pVal As Object) As NumericType

        Dim xRet As Object

        Try

            '' ���l���ł��Ȃ��ꍇ�́u0�v�ň���
            If IsNumeric(pVal) = False Then
                xRet = 0
            Else
                xRet = pVal
            End If

            '' NumericType�ɃL���X�g���Ė߂�
            Return CType(xRet, NumericType)

        Catch ex As Exception

            Throw New Exception("fToNumberEx:���l�ϊ��Ɏ��s���܂����" & vbCrLf & ex.Message)

        End Try

    End Function

    ''' <summary>fToText -- ������</summary>
    ''' <param name="pVal">�Ώەϐ�</param>
    ''' <returns>������</returns>
    ''' <remarks>pVal��÷�č��ډ����܂��NULL�w�肳�ꂽ�ꍇ�͋󔒂��ԋp����܂��</remarks>
    ''' <�쐬��></�쐬��>
    ''' <���l>DB���ڂ��e�L�X�g���ڂւ̓]�L�ȂǁCNull��������Ȃ����ڂւ̓]�L�Ɏg�p����</���l>
    <DebuggerStepThrough()>
    Public Shared Function fToText(ByVal pVal As Object) As String
        Dim ResultStr As String = String.Empty

        Try
            If pVal Is Nothing Then ''Nothing�ݒ�̏ꍇ�͋󔒕Ԃ�
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
    ''' fTrimSingleQuotes�̊֐�
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <history version="1" date="2017.09.21" name="PHONG-BTT">�V�K�쐬</history>
    Public Shared Function fTrimSingleQuotes(ByVal str As String) As String

        Return str.Trim("'")

    End Function

    ''' <summary>fToDate -- ���t���ډ�</summary>
    ''' <param name="pDate">�Ώەϐ�</param>
    ''' <param name="pErrDate">�G���[���̒l</param>
    ''' <param name="pRetType">�ԋp�l�̓��t����(�ȗ��� 1:���t+����,2:���t�̂�,3:���Ԃ̂�) ����l 1:���t+����</param>
    ''' <returns>���t</returns>
    ''' <remarks>pDate����t�ډ����܂�����t���ڊO���w�肳�ꂽ�ꍇ��,���Ұ�2��,�ȗ����ꂽ�ꍇ�ͼ��эŏ��l��ԋp���܂�</remarks>
    ''' <�쐬��>�Ĉ� �D��</�쐬��>
    ''' <���l></���l>
    <DebuggerStepThrough()>
    Public Shared Function fToDate(ByVal pDate As Object, Optional ByVal pErrDate As Date = #1/1/1000#, Optional ByVal pRetType As BrowsefToDate = BrowsefToDate.fToDate_DateTime) As Date
        Dim xDate As Date
        Dim ResultDate As Date

        Try
            '' ���t�����\���H
            If IsDate(pDate) Then
                xDate = CDate(pDate)
            Else
                xDate = pErrDate
            End If

            ''�ԋp�l�ʂɌ^�ϊ�
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

    ''' <summary>fToInt -- Int��</summary>
    ''' <param name="pVal">�ϊ��Ώےl</param>
    ''' <returns>Integer</returns>
    ''' <remarks>������fToNumberEx�֏����������n��</remarks>
    ''' <history version="1" date="2008.04.07" name="��t �F��">�V�K�쐬</history> 
    <DebuggerStepThrough()>
    Public Shared Function fToInt(ByVal pVal As Object) As Integer

        Return fToNumberEx(Of Integer)(pVal)

    End Function

    ''' <summary>fToDbl -- Double��</summary>
    ''' <param name="pVal">�ϊ��Ώےl</param>
    ''' <returns>Integer</returns>
    ''' <remarks>������fToNumberEx�֏����������n��</remarks>
    ''' <history version="1" date="2008.04.07" name="��t �F��">�V�K�쐬</history> 
    <DebuggerStepThrough()>
    Public Shared Function fToDbl(ByVal pVal As Object) As Double

        Return fToNumberEx(Of Double)(pVal)

    End Function

    ''' <summary>fToDec -- Decimal��</summary>
    ''' <param name="pVal">�ϊ��Ώےl</param>
    ''' <returns>Integer</returns>
    ''' <remarks>������fToNumberEx�֏����������n��</remarks>
    ''' <history version="1" date="2008.04.07" name="��t �F��">�V�K�쐬</history> 
    <DebuggerStepThrough()>
    Public Shared Function fToDec(ByVal pVal As Object) As Decimal

        Return fToNumberEx(pVal)

    End Function

    '��MidB
    ''' <summary>Mid�֐��̃o�C�g�ŁB�������ƈʒu���o�C�g���Ŏw�肵�ĕ������؂蔲���B</summary>
    ''' <param name="str">�Ώۂ̕�����</param>
    ''' <param name="Start">�؂蔲���J�n�ʒu�B�S�p�����𕪊�����悤�ʒu���w�肳�ꂽ�ꍇ�A�߂�l�̕�����̐擪�͈Ӗ��s���̔��p�����ƂȂ�B</param>
    ''' <param name="Length">�؂蔲��������̃o�C�g��</param>
    ''' <returns>�؂蔲���ꂽ������</returns>
    ''' <remarks>�Ō�̂P�o�C�g���S�p�����̔����ɂȂ�ꍇ�A���̂P�o�C�g�͖��������B</remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    ''' <history version="3" date="2009.06.01" name="���� ����">�؂蔲����������̃o�C�g���擾�s�����C��</history>
    Public Shared Function MidB(ByVal str As String, ByVal Start As Integer, Optional ByVal Length As Integer = 0) As String
        '���󕶎��ɑ΂��Ă͏�ɋ󕶎���Ԃ�

        If str Is Nothing Then
            Return String.Empty
        End If

        If str = "" Then
            Return ""
        End If

        '��Length�̃`�F�b�N

        'Length��0���AStart�ȍ~�̃o�C�g�����I�[�o�[����ꍇ��Start�ȍ~�̑S�o�C�g���w�肳�ꂽ���̂Ƃ݂Ȃ��B

        Dim RestLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(str) - Start + 1

        If Length = 0 OrElse Length > RestLength Then
            Length = RestLength
        End If

        '���؂蔲��

        Dim SJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
        Dim B() As Byte = CType(Array.CreateInstance(GetType(Byte), Length), Byte())

        Array.Copy(SJIS.GetBytes(str), Start - 1, B, 0, Length)

        Dim st1 As String = SJIS.GetString(B)

        '���؂蔲�������ʁA�Ō�̂P�o�C�g���S�p�����̔����������ꍇ�A���̔����͐؂�̂Ă�B

        ''2009.06.01 UPD START TAKASHIMA
        'Dim ResultLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(st1) - Start + 1
        Dim ResultLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(st1)
        ''2009.06.01 UPD END

        If Asc(Strings.Right(st1, 1)) = 0 Then
            'VB.NET2002,2003�̏ꍇ�A�Ō�̂P�o�C�g���S�p�̔����̎�
            Return st1.Substring(0, st1.Length - 1)
        ElseIf Length = ResultLength - 1 Then
            'VB2005�̏ꍇ�ōŌ�̂P�o�C�g���S�p�̔����̎�
            Return st1.Substring(0, st1.Length - 1)
        Else
            '���̑��̏ꍇ
            Return st1
        End If

    End Function

    ''' <summary>
    ''' LobDataDivision
    ''' </summary>
    ''' <param name="pLobStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    Public Shared Function LobDataDivision(ByVal pLobStr As String) As Array

        If pLobStr Is Nothing Then
            pLobStr = String.Empty
        End If

        Dim RetArr() As String
        Dim nByteLen As Decimal = LenB(pLobStr)

        '' �o�C�g�����S�O�O�O�o�C�g�𒴂��邩�`�F�b�N
        If nByteLen > 4000 Then
            '' �z��T�C�Y������
            Dim nArrCount As Decimal = fRound((nByteLen / 4000), 1, 1) - 1
            ReDim RetArr(nArrCount)

            '' �z��ɕ������Ēl��ݒ�
            Dim nCnt As Integer = 0
            Dim nTotalByte As Decimal = 0

            For i As Integer = nCnt To nArrCount

                '' �g�[�^���o�C�g���̏�����
                nTotalByte = 0

                '' ���݂̃g�[�^���o�C�g�����擾
                For j As Integer = 0 To nArrCount
                    If RetArr(j) <> Nothing Then
                        nTotalByte = nTotalByte + fToNumberEx(Of Decimal)(LenB(RetArr(j)))
                    End If
                Next

                '' �o�C�g����
                If i = nArrCount Then
                    RetArr(i) = MidB(pLobStr, (nTotalByte + 1), nByteLen - nTotalByte)
                Else
                    RetArr(i) = MidB(pLobStr, (nTotalByte + 1), 4000)
                End If
            Next

        Else
            '' �S�O�O�O�o�C�g�ȓ��̏ꍇ�͂��̂܂ܕԋp
            ReDim RetArr(0)
            RetArr(0) = pLobStr
        End If

        Return RetArr

    End Function

    ''' <summary>Camelize -- �A���_�[�X�R�A�q���̕�������L�������`���E�p�X�J���`���ɕϊ����܂�</summary>
    ''' <param name="str">������</param>
    ''' <param name="toUpper">�擪������啶���ɂ��邩�iTrue: ���� / False: ���Ȃ��j</param>
    ''' <returns>String</returns>
    ''' <remarks>�f�t�H���g�̓L�������`���i�擪�������j</remarks>
    ''' <history version="1" date="2009.08.27" name="�ɓ� ����">�V�K�쐬</history> 
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

    ''' <summary>Underscore -- �L�������`���E�p�X�J���`���̕�������A���_�[�X�R�A�q���ɕϊ����܂�</summary>
    ''' <param name="str">������</param>
    ''' <param name="toUpper">�啶���ɂ��邩�iTrue: �啶�� / False: �������j</param>
    ''' <returns>String</returns>
    ''' <remarks>�f�t�H���g�͑啶��</remarks>
    ''' <history version="1" date="2009.08.27" name="�ɓ� ����">�V�K�쐬</history> 
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

#Region " �G���[�ʒm "

    ''' <summary>ShowExceptionMessage -- ��O���烁�b�Z�[�W�{�b�N�X��\������</summary>
    ''' <param name="ex">��O�N���X</param>
    ''' <param name="MessageTitle">���b�Z�[�W�^�C�g��</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    ''' <history version="3" date="2009.04.24" name="��t �F��">�G���[���b�Z�[�W�̐�����ʃ��\�b�h�ɕ���</history>
    Public Shared Sub ShowExceptionMessage(ByVal ex As Exception, ByVal MessageTitle As String)
        '' =================================================================
        ''   [MessageTitle��\��]
        ''   ---------------------------------------------------------    
        ''
        ''   �y�G���[���e�z
        ''   [ex.Message��\��]    
        ''
        ''   �y�ǉ����z�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'''ex��InnerException���ݒ肳��Ă���ꍇ�ɕ\��
        ''   [ex.InnerException.Message��\��]    
        ''
        ''   �y�G���[���e�z
        ''   [ex.StackTrace��\��]    
        '' =================================================================

        If MessageTitle Is Nothing Then
            MessageTitle = "PROCES.S5"
        End If
        ' '' �ʒm���e��ҏW
        'Dim Message As New System.Text.StringBuilder
        'Message.AppendLine("�y�G���[���e�z")
        'Message.AppendLine(ex.Message)
        'Message.AppendLine()
        'If Not IsNothing(ex.InnerException) Then
        '    Message.AppendLine("�y�ǉ����z")
        '    Message.AppendLine(ex.InnerException.Message)
        '    Message.AppendLine()
        'End If
        'Message.AppendLine("�y�X�^�b�N�g���[�X�z")
        'Message.AppendLine(ex.StackTrace)

        '' ���b�Z�[�W�\��
        MessageBox.Show(CreateExceptionString(ex), MessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="lev"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history version="1" date="2009.04.24" name="��t �F��">�V�K�쐬</history>
    Public Shared Function CreateExceptionString(ByVal ex As Exception, Optional ByVal lev As Integer = 0) As String

        Dim Message As New System.Text.StringBuilder

        If lev.Equals(0) Then
            Message.AppendLine("�G���[���")
            Message.AppendLine()
        Else
            Message.AppendLine()
            Message.AppendLine("===================================================")
            Message.AppendFormat("�ǉ��G���[��� {0}", fToText(lev)) : Message.AppendLine()
            Message.AppendLine()
        End If

        Dim exceptionType As String
        If ex.GetType().Name = "ServiceCommonException" Then
            exceptionType = ex.ToString()
        Else
            exceptionType = ex.GetType().Name
        End If

        Message.AppendLine("�y�G���[�̎�ށz")
        Message.AppendLine()
        Message.AppendLine(exceptionType)
        Message.AppendLine()
        Message.AppendLine("�y���e�z")
        Message.AppendLine()
        Message.AppendLine(ex.Message)
        Message.AppendLine()
        Message.AppendLine("�y�X�^�b�N�g���[�X�z")
        Message.AppendLine()
        Message.AppendLine(ex.StackTrace)

        If Not IsNothing(ex.InnerException) Then
            Message.AppendLine(CreateExceptionString(ex.InnerException, lev + 1))
        End If

        Return Message.ToString

    End Function

    ''' <summary>ConvertException -- �e���O�N���X������O�N���X�ɕϊ�����</summary>
    ''' <param name="ex">��O�N���X</param>
    ''' <returns>����O�N���X</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertException(ByVal ex As Exception) As Exception

        Dim Result As Exception

        If ex.GetType Is GetType(SoapException) Then
            '' Web�T�[�r�X�Ŕ���������O�N���X�����̗�O�N���X�ɕϊ�
            Result = New Exception(CType(ex, SoapException).Detail.ChildNodes(0).InnerText, ex)

        Else
            '' ����܂ł̏����ŕϊ�����Ȃ���O�N���X�́A�P���ɗ�O�N���X�ɃL���X�g����
            Result = CType(ex, Exception)
        End If

        Return Result

    End Function

    ''' <summary>
    ''' WriteEventLog -- �C�x���g���O���o�͂���
    ''' </summary>
    ''' <param name="pEx">��O�N���X</param>
    ''' <param name="pMsg">�G���[���b�Z�[�W</param>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    ''' <history version="3" date="2011.12.14" name="���c �v�`">(SC�ŗ����錴���ɂȂ��Ă�����)EventLog�ւ̏������݃G���[�͉������悤���Ȃ��̂ŃX���[����</history>
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
            ''(SC�ŗ����錴���ɂȂ��Ă�����)EventLog�ւ̏������݃G���[�͉������悤���Ȃ��̂ŃX���[����
        End Try
        ''2011.12.14 UPD END

    End Sub

#End Region

#Region " �ʃA�v���N������ "

    ' ''' <summary>
    ' ''' ClicOnce �A�v���P�[�V�������N�����邽�߂� URL ���쐬���܂��B
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
    ' ''' ClicOnce �A�v���P�[�V�������N�����邽�߂� IE �̃v���Z�X�����쐬���܂��B
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
    ''' <history version="1" date="2009.07.22" name="��t �F��">�I�[�o�[���[�h�����i���������̂܂܃Z�b�g����j</history>
    Public Shared Function CreateProcessInfo(ByVal arg As String, ByVal flg As ProcessWindowStyle) As System.Diagnostics.ProcessStartInfo

        Dim info As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo

        ''�N�����̍쐬�i����V����IE����ċN������悤�ɕύX�j
        info.FileName = "IExplore.exe"                                      '' IE�v���Z�X���i�Œ�j
        info.Arguments = arg                                                '' �������_�C���N�g�ɃZ�b�g
        info.UseShellExecute = True                                         '' �V�F���N��
        info.CreateNoWindow = False                                         '' �V����Window���쐬
        info.WindowStyle = flg                                              '' �������_�C���N�g�ɃZ�b�g

        Return info

    End Function

#End Region

#Region " ��ԃ��x���ύX���� "

    ''' <summary>
    ''' SetLabelConfig
    ''' </summary>
    ''' <param name="LabelControl"></param>
    ''' <param name="Index"></param>
    ''' <remarks></remarks>
    ''' <history version="1" date="" name="">���ŕs��</history>
    ''' <history version="2" date="2009.03.25" name="��t �F��">��������ǉ�</history>
    Public Shared Sub SetLabelConfig(ByVal LabelControl As Label, ByVal Index As BrouseConfig)

        If LabelControl Is Nothing Then
            Return
        End If

        Select Case Index
            Case BrouseConfig.NoData
                '���o�^����
                LabelControl.Text = "�R�[�h������"
                LabelControl.ForeColor = System.Drawing.Color.Khaki
                LabelControl.BackColor = System.Drawing.Color.Olive
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.NewData
                '�V�K�o�^����(���͊J�n�O)
                LabelControl.Text = "���o�^�R�[�h"
                LabelControl.ForeColor = System.Drawing.Color.PowderBlue
                LabelControl.BackColor = System.Drawing.Color.Navy
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.NewDataInput
                '�V�K�o�^����(���͊J�n��)
                LabelControl.Text = "�V�K���͒�"
                LabelControl.ForeColor = System.Drawing.Color.Navy
                LabelControl.BackColor = System.Drawing.Color.PowderBlue
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.UpdData
                '�C���o�^����(���͊J�n�O)
                LabelControl.Text = "�Ɖ�\����"
                LabelControl.ForeColor = System.Drawing.Color.PaleGreen
                LabelControl.BackColor = System.Drawing.Color.DarkGreen
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.UpdDataInput
                '�C���o�^����(���͊J�n��)
                LabelControl.Text = "�C�����͒�"
                LabelControl.ForeColor = System.Drawing.Color.DarkGreen
                LabelControl.BackColor = System.Drawing.Color.PaleGreen
                LabelControl.BorderStyle = BorderStyle.FixedSingle
            Case BrouseConfig.NoSec
                '������������
                LabelControl.Text = "�������R�[�h"
                LabelControl.ForeColor = System.Drawing.Color.Red
                LabelControl.BackColor = System.Drawing.Color.Black
                LabelControl.BorderStyle = BorderStyle.FixedSingle
                ''2008.07.01 ADD START T.Nara
            Case BrouseConfig.Running
                '�f�[�^�X�V���s������
                LabelControl.Text = "�X�V��"
                LabelControl.ForeColor = System.Drawing.Color.LightPink
                LabelControl.BackColor = System.Drawing.Color.Maroon
                LabelControl.BorderStyle = BorderStyle.FixedSingle
                ''2008.07.01 ADD END
            Case BrouseConfig.Initialize
                '������(���x���̔�\��)
                LabelControl.Text = ""
                LabelControl.ForeColor = LabelControl.Parent.BackColor
                LabelControl.BackColor = LabelControl.Parent.BackColor
                LabelControl.BorderStyle = BorderStyle.None
        End Select

    End Sub

#End Region

#Region " ���匠���֘A "

#Region " ���匠�� "
    ''' <summary>
    ''' GetWhereStrBmnSecurity -- ���匠�����l�����邽�߂�Where����擾
    ''' </summary>
    ''' <param name="checkType">CheckBmnSecType�񋓑�</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where��</returns>
    ''' <remarks>SearchPanel����̌ďo���l������CommonMethod�Ɏ���
    '''          Form�A�v���Ŏg�p����ꍇ��BaseForm�̓���Method���ďo������</remarks>
    ''' <history version="1" date="2010.05.13" name="���� ����">�V�K�쐬</history>
    Public Overloads Shared Function GetWhereStrBmnSecurity(ByVal checkType As CheckBmnSecType,
                                                            Optional ByVal userNo As Integer = 0) As String
        ' �ʏ��CODE�Ō����`�F�b�N
        Return GetWhereStrBmnSecurity("CODE", checkType, userNo)
    End Function

    ''' <summary>
    ''' GetWhereStrBmnSecurity -- ���匠�����l�����邽�߂�Where����擾
    ''' </summary>
    ''' <param name="field">Field�� in (�����}�X�^.CODE)</param>
    ''' <param name="checkType">CheckBmnSecType�񋓑�</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where��</returns>
    ''' <remarks>SearchPanel����̌ďo���l������CommonMethod�Ɏ���
    '''          Form�A�v���Ŏg�p����ꍇ��BaseForm�̓���Method���ďo������</remarks>
    ''' <history version="1" date="2010.05.13" name="���� ����">�V�K�쐬</history>
    Public Overloads Shared Function GetWhereStrBmnSecurity(ByVal field As String,
                                                            ByVal checkType As CheckBmnSecType,
                                                            Optional ByVal userNo As Integer = 0) As String
        Dim xSqlTxt As String

        ' Where���SQL����
        xSqlTxt = String.Format("		{0} IN ( ", field)  ' �w�肳�ꂽField���g�p
        xSqlTxt &= vbCrLf & "				SELECT "
        xSqlTxt &= vbCrLf & "						CODE "
        xSqlTxt &= vbCrLf & "				FROM "
        xSqlTxt &= vbCrLf & "						{0} "
        xSqlTxt &= vbCrLf & "				WHERE "
        xSqlTxt &= vbCrLf & "						USER_NO	=	{1} "
        xSqlTxt &= vbCrLf & "		) "

        ' ��p�̕ϊ����������ĕԋp
        Return FormatStrBmnSecurity(xSqlTxt, checkType, userNo)
    End Function
#End Region
#Region " ����K�w���� "
    ''' <summary>
    ''' GetWhereStrBmnKaisouSecurity -- ���匠�����l�����邽�߂�Where����擾
    ''' </summary>
    ''' <param name="checkType">CheckBmnSecType�񋓑�</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where��</returns>
    ''' <remarks>SearchPanel����̌ďo���l������CommonMethod�Ɏ���
    '''          Form�A�v���Ŏg�p����ꍇ��BaseForm�̓���Method���ďo������</remarks>
    ''' <history version="1" date="2010.05.13" name="���� ����">�V�K�쐬</history>
    Public Overloads Shared Function GetWhereStrBmnKaisouSecurity(ByVal checkType As CheckBmnSecType,
                                                                  Optional ByVal userNo As Integer = 0) As String
        ' �ʏ��CODE�Ō����`�F�b�N
        Return GetWhereStrBmnKaisouSecurity("CODE", checkType, userNo)
    End Function

    ''' <summary>
    ''' GetWhereStrBmnSecurity -- ���匠�����l�����邽�߂�Where����擾
    ''' </summary>
    ''' <param name="field">Field�� in (�����}�X�^.CODE)</param>
    ''' <param name="checkType">CheckBmnSecType�񋓑�</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>Where��</returns>
    ''' <remarks>SearchPanel����̌ďo���l������CommonMethod�Ɏ���
    '''          Form�A�v���Ŏg�p����ꍇ��BaseForm�̓���Method���ďo������</remarks>
    ''' <history version="1" date="2010.05.13" name="���� ����">�V�K�쐬</history>
    Public Overloads Shared Function GetWhereStrBmnKaisouSecurity(ByVal field As String,
                                                                  ByVal checkType As CheckBmnSecType,
                                                                  Optional ByVal userNo As Integer = 0) As String
        Dim xSqlTxt As String

        ' Where���SQL����
        xSqlTxt = String.Format("		{0} IN ( ", field)  ' �w�肳�ꂽField���g�p
        xSqlTxt &= vbCrLf & "				SELECT "
        xSqlTxt &= vbCrLf & "						KAISOU "
        xSqlTxt &= vbCrLf & "				FROM "
        xSqlTxt &= vbCrLf & "						{0} "
        xSqlTxt &= vbCrLf & "				WHERE "
        xSqlTxt &= vbCrLf & "						USER_NO	=	{1} "
        xSqlTxt &= vbCrLf & "				GROUP BY "
        xSqlTxt &= vbCrLf & "						KAISOU "
        xSqlTxt &= vbCrLf & "		) "

        ' ��p�̕ϊ����������ĕԋp
        Return FormatStrBmnSecurity(xSqlTxt, checkType, userNo)
    End Function
#End Region

    ''' <summary>
    ''' FormatStrBmnSecurity(Private) -- ���匠���l���p��String.Format
    ''' </summary>
    ''' <param name="xSqlTxt">�ϊ����ڂ��܂�String({0}:�����}�X�^, {1}:USER_NO)</param>
    ''' <param name="checkType">CheckBmnSecType�񋓑�</param>
    ''' <param name="userNo">USER_NO</param>
    ''' <returns>�ϊ����String</returns>
    ''' <remarks>Private</remarks>
    ''' <history version="1" date="2010.05.13" name="���� ����">�V�K�쐬</history>
    Private Shared Function FormatStrBmnSecurity(ByVal xSqlTxt As String, ByVal checkType As CheckBmnSecType, ByVal userNo As Integer) As String

        Dim xBmnSecMaster As String
        Select Case checkType
            Case CheckBmnSecType.Sansyo
                xBmnSecMaster = "S_BMN_UW" ' �Q�ƌ���
            Case CheckBmnSecType.Input
                xBmnSecMaster = "I_BMN_UW" ' ���͌���
            Case Else
                xBmnSecMaster = "S_BMN_UW" ' �Q�ƌ���(Else)
        End Select

        Dim oUserNo As Object
        If userNo = 0 Then
            oUserNo = "@SYSUSERNO@" ' ServiceCommon�Œu�������USER_NO�̑��Text
        Else
            oUserNo = fSqlStr(userNo, BrowsefSqlStr.fSqlStr_Integer)
        End If

        ' �ϊ����ĕԋp
        Return String.Format(xSqlTxt, xBmnSecMaster, oUserNo)
    End Function

#End Region

End Class



