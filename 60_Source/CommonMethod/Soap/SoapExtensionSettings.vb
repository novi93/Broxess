Imports System.Configuration
Imports System.Web.Services.Configuration

Namespace Soap

    ''' <summary>
    ''' �N���C�A���g�p app.config�ݒ�iSoapExtention�j
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history version="5.0.0.0" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
    Public Class SoapExtensionSettings

#Region "  �萔�錾  "

        ' app.config �ǎ�p�v���p�e�B�l
        Private Const SOAP_EXTENSION_SECTION_GROUP As String = "system.web"
        Private Const SOAP_EXTENSION_SECTION As String = "webServices"
        Private Const SOAP_EXTENSION_PROPERTIES As String = "soapExtensionTypes"

#End Region

#Region "  Shared�ϐ�   "

        ' app.config
        Private Shared _appConfig As Configuration = _
                        ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        ' app.config SoapExtension�ݒ�I�u�W�F�N�g
        Private Shared _soapExtensionType As PropertyInformation = _
                        _appConfig.GetSectionGroup(SOAP_EXTENSION_SECTION_GROUP) _
                            .Sections(SOAP_EXTENSION_SECTION) _
                            .ElementInformation.Properties(SOAP_EXTENSION_PROPERTIES)

#End Region

#Region "  Shared Method  "

        ''' <summary>
        ''' app.cofing�ɁAsoapExtensionTypes���ݒ肳��Ă��邩���肵�܂��B
        ''' </summary>
        ''' <returns>soapExtensionTypes���ݒ肳��Ă��邩</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        Public Shared Function IsCompressionSoapExtensionSetteing() As Boolean

            Dim result As Boolean = False

            Dim collection As SoapExtensionTypeElementCollection = _
                    DirectCast(_soapExtensionType.Value, SoapExtensionTypeElementCollection)

            For Each element As SoapExtensionTypeElement In collection
                If element.Type Is GetType(CompressionSoapExtension) Then
                    result = True
                End If
            Next

            Return result

        End Function


        ''' <summary>
        ''' app.config�ɁAsoapExtention��ݒ肵�܂��B
        ''' </summary>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="�ɓ� ����">�V�K�쐬</history>
        Public Shared Sub CompressionSoapExtensionSetting()

            If IsCompressionSoapExtensionSetteing() = False Then

                Dim exCollection As New SoapExtensionTypeElementCollection
                Dim ex As New SoapExtensionTypeElement
                ex.Type = GetType(CompressionSoapExtension)
                ex.Priority = 1
                ex.Group = Web.Services.Configuration.PriorityGroup.Low
                exCollection.Add(ex)
                _soapExtensionType.Value = exCollection

                _appConfig.Save()

            End If

        End Sub

#End Region

    End Class

End Namespace
