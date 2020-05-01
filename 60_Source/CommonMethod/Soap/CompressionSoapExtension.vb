Imports System.IO
Imports System.Web.Services.Protocols
Imports System.IO.Compression
Imports System.Xml
Imports System.Text
Imports System.Collections.Generic

Namespace Soap

    ''' <summary>
    ''' Soap�ʐM���kExtention
    ''' </summary>
    ''' <remarks>���k�ۃt���O��ON�̏ꍇ�̂݁ADeflate���k�������܂��B</remarks>
    ''' <history version="5.0.0.0" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
    ''' <history version="5.0.0.0" date="2010.01.04" name="��t �F��">���k�t���O���������ύX�i���ʌ݊��s�\�j</history>
    Public Class CompressionSoapExtension
        Inherits SoapExtension

#Region "  �萔�錾  "

        '' 2010.01.04 DEL START
        '' SoapHeader�I���^�O
        'Private Const SOAP_HEADER_END As String = "</soap:Header>"
        '' SoapBody�J�n�^�O
        'Private Const SOAP_BODY_START As String = "<soap:Body>"
        '' ProcessSoapHeader�I���^�O
        'Private Const PROCESS_SOAP_HEADER_END As String = "</ProcessSoapHeader>"
        ' ���k�ۃt���O
        'Private Const IS_COMPRESSION_TRUE As String = "IsCompression=""true"""
        'Private Const IS_COMPRESSION_FALSE As String = "IsCompression=""false"""

        '' �I������^�O�����p�^�[��
        'Private Const REG_PATURN As String = "(</.*:Header>)|(<[^/].*:Body>)"      '�����̃p�^�[���g���ƁA���x
        'Private Const REG_PAT_HD_ST As String = "<[^/].*:Header"
        'Private Const REG_PAT_HD_ED As String = "</.*:Header>"
        'Private Const REG_PAT_BD_ST As String = "<[^/].*:Body>"
        '' 2010.01.04 DEL END

        '' ���M�����k����p
        Private IsComp As Boolean
        Private Const TYPE_OF_COMP As String = "deflate"

#End Region

#Region "  �C���X�^���X�ϐ�  "

        'BeforeDeserialize�ŃV���A�������ꂽSOAP���b�Z�[�W��ێ����郁���o�ϐ�
        Private _oldStream As Stream
        'AfterSerialize�ŃV���A�������ꂽSOAP���b�Z�[�W��ێ����郁���o�ϐ�
        Private _newStream As Stream

#End Region

#Region "  Public Method  "

        ''' <summary>
        ''' �J�X�^��������K�p����悤�Ɋg���@�\���w�肵���ꍇ���s����܂�
        ''' �߂�l�ɂ�SOAP�g���@�\�ŃL���b�V���O�������l��Ԃ��܂� 
        ''' </summary>
        ''' <param name="methodInfo">LogicalMethodInfo</param>
        ''' <param name="attribute">SoapExtensionAttribute</param>
        ''' <returns>SOAP�g���@�\�ŃL���b�V���O�������l</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        Public Overloads Overrides Function GetInitializer( _
                        ByVal methodInfo As System.Web.Services.Protocols.LogicalMethodInfo, _
                        ByVal attribute As System.Web.Services.Protocols.SoapExtensionAttribute) As Object
            Return Nothing
        End Function

        ''' <summary>
        ''' web.config�\���t�@�C���܂���app.config�t�@�C���ɎQ�Ƃ�ǉ������ꍇ���s����܂�
        ''' �߂�l�ɂ�SOAP�g���@�\�ŃL���b�V���O�������l��Ԃ��܂� 
        ''' </summary>
        ''' <param name="serviceType"></param>
        ''' <returns>SOAP�g���@�\�ŃL���b�V���O�������l</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        Public Overloads Overrides Function GetInitializer(ByVal serviceType As System.Type) As Object
            Return Nothing
        End Function

        ''' <summary>
        ''' SOAP�@�\�g���̏��������Ɉ�x�̂ݎ��s����܂�
        ''' GetInitializer���\�b�h�ł̖߂�l���p�����[�^�Ƃ��ēn����܂�
        ''' </summary>
        ''' <param name="initializer"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        Public Overloads Overrides Sub Initialize(ByVal initializer As Object)
        End Sub

        ''' <summary>
        ''' �����œn���ꂽSOAP���b�Z�[�W�ւ̎Q�Ƃ������o�ϐ��ɐݒ肵�܂�
        ''' �߂�l��SOAP�@�\�g���ŗ��p�����߂�l�ւ̎Q�ƂƂȂ�܂�
        ''' </summary>
        ''' <param name="stream"></param>
        ''' <returns>SOAP�@�\�g���ŗ��p�����߂�l�ւ̎Q��</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        Public Overloads Overrides Function ChainStream(ByVal stream As Stream) As Stream
            Me._oldStream = stream
            Me._newStream = New MemoryStream()

            Return Me._newStream
        End Function

        ''' <summary>
        ''' SOAP�g���@�\�̂��ׂĂ�SoapMessageStage�i�K�Ŏ��s����܂�
        ''' </summary>
        ''' <param name="message"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        ''' <history version="2" date="2010.01.04" name="��t �F��">�V���A�����O�ɃR���e���c�G���R�[�h���w�肷��悤�ɕύX�A�t�V���A�����O�ɃR���e���c�G���R�[�h���爳�k����Ă��邩�ۂ����擾</history>
        Public Overloads Overrides Sub ProcessMessage(ByVal message As System.Web.Services.Protocols.SoapMessage)

            Select Case message.Stage
                Case SoapMessageStage.BeforeSerialize

                    '' SOAP���k�t���O�̓ǂݏo��
                    For Each sh As SoapHeader In message.Headers
                        '' ProcessSoapHeader�ȊO�͖���
                        If Not sh.GetType.Name.Equals("ProcessSoapHeader") Then Continue For

                        '' �v���p�e�B�����o�m�F
                        For Each p As System.Reflection.PropertyInfo In sh.GetType.GetProperties()
                            '' IsCompression�ȊO�͖���
                            If Not p.Name.Equals("IsCompression") Then Continue For

                            '' IsCompression���擾
                            IsComp = DirectCast(p.GetValue(sh, Nothing), Boolean)
                            If IsComp Then
                                '' �R���e���c�̃G���R�[�h�Ɉ��k���w��
                                message.ContentEncoding = TYPE_OF_COMP
                            End If
                            Exit Select
                        Next
                    Next

                Case SoapMessageStage.AfterSerialize

                    _newStream.Position = 0
                    '�V���A���C�Y��Ɉ��k���܂��iSOAP�S�����Ώہj
                    Me.CompressStream(_newStream, _oldStream)

                Case SoapMessageStage.BeforeDeserialize

                    '' �R���e���c�G���R�[�h�����k���ۂ�
                    IsComp = CommonMethod.fToText(message.ContentEncoding).Equals(TYPE_OF_COMP)

                    '�f�V���A���C�Y�O�ɉ𓀂��܂��iSOAP�S�����Ώہj
                    Me.DecompressStream(_oldStream, _newStream)
                    _newStream.Position = 0

            End Select

        End Sub

#End Region

#Region "  Private Method  "

        ''' <summary>
        ''' �X�g���[�����k����
        ''' </summary>
        ''' <param name="from"></param>
        ''' <param name="To"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        ''' <history version="2" date="2010.01.04" name="��t �F��">���k�t���O���N���X�ϐ�����擾����悤�ɕύX</history>
        Private Sub CompressStream(ByVal from As Stream, ByVal [To] As Stream)

            '' SoapHeader���A���k�ۃt���O���擾
            'Dim isCompres As Boolean = CompressionSoapExtension.IsCompression(from, [To])
            Dim isCompres As Boolean = IsComp

            Dim buffer(1023) As Byte
            Dim readSize As Integer

            If isCompres Then

                ' ���k����
                Using compressedStream As New DeflateStream([To], CompressionMode.Compress, True)

                    While True
                        readSize = from.Read(buffer, 0, buffer.Length)
                        If readSize = 0 Then
                            Exit While
                        End If
                        compressedStream.Write(buffer, 0, readSize)
                    End While
                    compressedStream.Close()
                    [To].Flush()
                End Using

            Else
                ' ���k���Ȃ�����
                While True
                    readSize = from.Read(buffer, 0, buffer.Length)
                    If readSize = 0 Then
                        Exit While
                    End If
                    [To].Write(buffer, 0, readSize)
                End While
                [To].Flush()
            End If

        End Sub

        ''' <summary>
        ''' �X�g���[���𓀏���
        ''' </summary>
        ''' <param name="from"></param>
        ''' <param name="To"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        ''' <history version="2" date="2010.01.04" name="��t �F��">���k�t���O���N���X�ϐ�����擾����悤�ɕύX</history>
        Private Sub DecompressStream(ByVal from As Stream, ByVal [To] As Stream)

            ' SoapHeader���A���k�ۃt���O���擾
            'Dim isCompres As Boolean = CompressionSoapExtension.IsCompression(from, [To])
            Dim isCompres As Boolean = IsComp

            Dim buffer(1023) As Byte
            Dim readSize As Integer
            If isCompres Then
                ' �𓀏���
                Using compressedStream As DeflateStream = New DeflateStream(from, CompressionMode.Decompress)

                    While True
                        readSize = compressedStream.Read(buffer, 0, buffer.Length)
                        If readSize = 0 Then
                            Exit While
                        End If
                        [To].Write(buffer, 0, readSize)
                    End While
                    [To].Flush()
                End Using
            Else
                ' �������Ȃ�����
                While True
                    readSize = from.Read(buffer, 0, buffer.Length)
                    If readSize = 0 Then
                        Exit While
                    End If
                    [To].Write(buffer, 0, readSize)
                End While
                [To].Flush()
            End If

        End Sub

        '' 2010.01.04 DEL START
        '''' <summary>
        '''' SoapHeader���A���k�ۃt���O���擾
        '''' </summary>
        '''' <param name="from"></param>
        '''' <param name="To"></param>
        '''' <returns>���k���邩�ǂ����̃t���O</returns>
        '''' <remarks></remarks>
        'Private Shared Function IsCompression(ByVal from As Stream, ByVal [To] As Stream) As Boolean

        '    Dim result As Boolean = False
        '    Dim strBuf As New StringBuilder
        '    Dim byteBuf As New List(Of Byte)
        '    Dim readInt As Integer

        '    'Dim regex As New System.Text.RegularExpressions.Regex(REG_PATURN)
        '    Dim regHead As New System.Text.RegularExpressions.Regex(REG_PAT_HD_ST, RegularExpressions.RegexOptions.Compiled)
        '    Dim regTarget As New System.Text.RegularExpressions.Regex(REG_PAT_BD_ST, RegularExpressions.RegexOptions.Compiled)

        '    Dim UseHeader As Boolean = False

        '    While True

        '        readInt = from.ReadByte()
        '        If readInt = -1 Then
        '            ' �ǂݍ��݂��Ō�܂œ��B�����ꍇ�A�����I��
        '            [To].Write(byteBuf.ToArray, 0, byteBuf.ToArray.Length)
        '            Exit While
        '        End If
        '        If Not IsNothing(Convert.ToChar(readInt)) Then
        '            strBuf.Append(Convert.ToChar(readInt))
        '            Debug.Print(Convert.ToChar(readInt))
        '        End If
        '        If Not IsNothing(CType(readInt, Byte)) Then
        '            byteBuf.Add(CType(readInt, Byte))
        '        End If

        '        If byteBuf.Count <= 150 Then
        '            '' 150�o�C�g�܂ł͖������œǂݍ���
        '            Continue While
        '        ElseIf UseHeader And byteBuf.Count <= 500 Then
        '            '' �w�b�_�g�p���ʌ�A500�o�C�g�܂ł͖������œǂݍ���
        '            Continue While
        '        End If

        '        '' Header���܂��������Ă��Ȃ�
        '        If Not UseHeader Then
        '            '' Header�����邩�H
        '            UseHeader = regHead.IsMatch(strBuf.ToString)
        '            If UseHeader Then
        '                '' ���������ꍇ�́AHeader�̏I���^�O���^�[�Q�b�g�ɐ؂�ւ���
        '                regTarget = New System.Text.RegularExpressions.Regex(REG_PAT_HD_ED, RegularExpressions.RegexOptions.Compiled)
        '                Continue While
        '            End If
        '        End If

        '        'If strBuf.ToString.Contains(PROCESS_SOAP_HEADER_END) OrElse strBuf.ToString.Contains(SOAP_HEADER_END) OrElse strBuf.ToString.Contains(SOAP_BODY_START) Then
        '        'If IsCompletedReadHeader(regex, strBuf) Then
        '        If regTarget.IsMatch(strBuf.ToString) Then
        '            ' SoapHeader������Ǎ�����
        '            If strBuf.ToString.Contains(IS_COMPRESSION_TRUE) Then
        '                ' "SoapHeader�ɁAIsCompression="true"���܂ޏꍇ
        '                ' ���k����
        '                result = True
        '            End If
        '            [To].Write(byteBuf.ToArray, 0, byteBuf.ToArray.Length)
        '            Exit While
        '        End If

        '    End While

        '    Return result

        'End Function
        '' 2010.01.04 DEL END

#End Region

    End Class

End Namespace