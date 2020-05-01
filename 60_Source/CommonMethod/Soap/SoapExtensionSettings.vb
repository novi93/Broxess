Imports System.Configuration
Imports System.Web.Services.Configuration

Namespace Soap

    ''' <summary>
    ''' クライアント用 app.config設定（SoapExtention）
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history version="5.0.0.0" date="2009.10.22" name="伊藤 匡明">新規作成</history>
    Public Class SoapExtensionSettings

#Region "  定数宣言  "

        ' app.config 読取用プロパティ値
        Private Const SOAP_EXTENSION_SECTION_GROUP As String = "system.web"
        Private Const SOAP_EXTENSION_SECTION As String = "webServices"
        Private Const SOAP_EXTENSION_PROPERTIES As String = "soapExtensionTypes"

#End Region

#Region "  Shared変数   "

        ' app.config
        Private Shared _appConfig As Configuration = _
                        ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        ' app.config SoapExtension設定オブジェクト
        Private Shared _soapExtensionType As PropertyInformation = _
                        _appConfig.GetSectionGroup(SOAP_EXTENSION_SECTION_GROUP) _
                            .Sections(SOAP_EXTENSION_SECTION) _
                            .ElementInformation.Properties(SOAP_EXTENSION_PROPERTIES)

#End Region

#Region "  Shared Method  "

        ''' <summary>
        ''' app.cofingに、soapExtensionTypesが設定されているか判定します。
        ''' </summary>
        ''' <returns>soapExtensionTypesが設定されているか</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
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
        ''' app.configに、soapExtentionを設定します。
        ''' </summary>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
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
