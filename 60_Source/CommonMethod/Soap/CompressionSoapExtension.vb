Imports System.IO
Imports System.Web.Services.Protocols
Imports System.IO.Compression
Imports System.Xml
Imports System.Text
Imports System.Collections.Generic

Namespace Soap

    ''' <summary>
    ''' Soap通信圧縮Extention
    ''' </summary>
    ''' <remarks>圧縮可否フラグがONの場合のみ、Deflate圧縮をかけます。</remarks>
    ''' <history version="5.0.0.0" date="2009.10.22" name="伊藤 匡明">新規作成</history>
    ''' <history version="5.0.0.0" date="2010.01.04" name="千葉 友則">圧縮フラグ判定方式を変更（下位互換不能）</history>
    Public Class CompressionSoapExtension
        Inherits SoapExtension

#Region "  定数宣言  "

        '' 2010.01.04 DEL START
        '' SoapHeader終了タグ
        'Private Const SOAP_HEADER_END As String = "</soap:Header>"
        '' SoapBody開始タグ
        'Private Const SOAP_BODY_START As String = "<soap:Body>"
        '' ProcessSoapHeader終了タグ
        'Private Const PROCESS_SOAP_HEADER_END As String = "</ProcessSoapHeader>"
        ' 圧縮可否フラグ
        'Private Const IS_COMPRESSION_TRUE As String = "IsCompression=""true"""
        'Private Const IS_COMPRESSION_FALSE As String = "IsCompression=""false"""

        '' 終了判定タグ検索パターン
        'Private Const REG_PATURN As String = "(</.*:Header>)|(<[^/].*:Body>)"      '←このパターン使うと、激遅
        'Private Const REG_PAT_HD_ST As String = "<[^/].*:Header"
        'Private Const REG_PAT_HD_ED As String = "</.*:Header>"
        'Private Const REG_PAT_BD_ST As String = "<[^/].*:Body>"
        '' 2010.01.04 DEL END

        '' 送信側圧縮判定用
        Private IsComp As Boolean
        Private Const TYPE_OF_COMP As String = "deflate"

#End Region

#Region "  インスタンス変数  "

        'BeforeDeserializeでシリアル化されたSOAPメッセージを保持するメンバ変数
        Private _oldStream As Stream
        'AfterSerializeでシリアル化されたSOAPメッセージを保持するメンバ変数
        Private _newStream As Stream

#End Region

#Region "  Public Method  "

        ''' <summary>
        ''' カスタム属性を適用するように拡張機能を指定した場合実行されます
        ''' 戻り値にはSOAP拡張機能でキャッシングしたい値を返します 
        ''' </summary>
        ''' <param name="methodInfo">LogicalMethodInfo</param>
        ''' <param name="attribute">SoapExtensionAttribute</param>
        ''' <returns>SOAP拡張機能でキャッシングしたい値</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        Public Overloads Overrides Function GetInitializer( _
                        ByVal methodInfo As System.Web.Services.Protocols.LogicalMethodInfo, _
                        ByVal attribute As System.Web.Services.Protocols.SoapExtensionAttribute) As Object
            Return Nothing
        End Function

        ''' <summary>
        ''' web.config構成ファイルまたはapp.configファイルに参照を追加した場合実行されます
        ''' 戻り値にはSOAP拡張機能でキャッシングしたい値を返します 
        ''' </summary>
        ''' <param name="serviceType"></param>
        ''' <returns>SOAP拡張機能でキャッシングしたい値</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        Public Overloads Overrides Function GetInitializer(ByVal serviceType As System.Type) As Object
            Return Nothing
        End Function

        ''' <summary>
        ''' SOAP機能拡張の初期化時に一度のみ実行されます
        ''' GetInitializerメソッドでの戻り値がパラメータとして渡されます
        ''' </summary>
        ''' <param name="initializer"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        Public Overloads Overrides Sub Initialize(ByVal initializer As Object)
        End Sub

        ''' <summary>
        ''' 引数で渡されたSOAPメッセージへの参照をメンバ変数に設定します
        ''' 戻り値はSOAP機能拡張で利用される戻り値への参照となります
        ''' </summary>
        ''' <param name="stream"></param>
        ''' <returns>SOAP機能拡張で利用される戻り値への参照</returns>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        Public Overloads Overrides Function ChainStream(ByVal stream As Stream) As Stream
            Me._oldStream = stream
            Me._newStream = New MemoryStream()

            Return Me._newStream
        End Function

        ''' <summary>
        ''' SOAP拡張機能のすべてのSoapMessageStage段階で実行されます
        ''' </summary>
        ''' <param name="message"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        ''' <history version="2" date="2010.01.04" name="千葉 友則">シリアル化前にコンテンツエンコードを指定するように変更、逆シリアル化前にコンテンツエンコードから圧縮されているか否かを取得</history>
        Public Overloads Overrides Sub ProcessMessage(ByVal message As System.Web.Services.Protocols.SoapMessage)

            Select Case message.Stage
                Case SoapMessageStage.BeforeSerialize

                    '' SOAP圧縮フラグの読み出し
                    For Each sh As SoapHeader In message.Headers
                        '' ProcessSoapHeader以外は無視
                        If Not sh.GetType.Name.Equals("ProcessSoapHeader") Then Continue For

                        '' プロパティメンバ確認
                        For Each p As System.Reflection.PropertyInfo In sh.GetType.GetProperties()
                            '' IsCompression以外は無視
                            If Not p.Name.Equals("IsCompression") Then Continue For

                            '' IsCompressionを取得
                            IsComp = DirectCast(p.GetValue(sh, Nothing), Boolean)
                            If IsComp Then
                                '' コンテンツのエンコードに圧縮を指定
                                message.ContentEncoding = TYPE_OF_COMP
                            End If
                            Exit Select
                        Next
                    Next

                Case SoapMessageStage.AfterSerialize

                    _newStream.Position = 0
                    'シリアライズ後に圧縮します（SOAP全文が対象）
                    Me.CompressStream(_newStream, _oldStream)

                Case SoapMessageStage.BeforeDeserialize

                    '' コンテンツエンコードが圧縮か否か
                    IsComp = CommonMethod.fToText(message.ContentEncoding).Equals(TYPE_OF_COMP)

                    'デシリアライズ前に解凍します（SOAP全文が対象）
                    Me.DecompressStream(_oldStream, _newStream)
                    _newStream.Position = 0

            End Select

        End Sub

#End Region

#Region "  Private Method  "

        ''' <summary>
        ''' ストリーム圧縮処理
        ''' </summary>
        ''' <param name="from"></param>
        ''' <param name="To"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        ''' <history version="2" date="2010.01.04" name="千葉 友則">圧縮フラグをクラス変数から取得するように変更</history>
        Private Sub CompressStream(ByVal from As Stream, ByVal [To] As Stream)

            '' SoapHeaderより、圧縮可否フラグを取得
            'Dim isCompres As Boolean = CompressionSoapExtension.IsCompression(from, [To])
            Dim isCompres As Boolean = IsComp

            Dim buffer(1023) As Byte
            Dim readSize As Integer

            If isCompres Then

                ' 圧縮処理
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
                ' 圧縮しない処理
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
        ''' ストリーム解凍処理
        ''' </summary>
        ''' <param name="from"></param>
        ''' <param name="To"></param>
        ''' <remarks></remarks>
        ''' <history version="1" date="2009.10.22" name="伊藤 匡明">新規作成</history>
        ''' <history version="2" date="2010.01.04" name="千葉 友則">圧縮フラグをクラス変数から取得するように変更</history>
        Private Sub DecompressStream(ByVal from As Stream, ByVal [To] As Stream)

            ' SoapHeaderより、圧縮可否フラグを取得
            'Dim isCompres As Boolean = CompressionSoapExtension.IsCompression(from, [To])
            Dim isCompres As Boolean = IsComp

            Dim buffer(1023) As Byte
            Dim readSize As Integer
            If isCompres Then
                ' 解凍処理
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
                ' 解消しない処理
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
        '''' SoapHeaderより、圧縮可否フラグを取得
        '''' </summary>
        '''' <param name="from"></param>
        '''' <param name="To"></param>
        '''' <returns>圧縮するかどうかのフラグ</returns>
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
        '            ' 読み込みが最後まで到達した場合、処理終了
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
        '            '' 150バイトまでは無条件で読み込み
        '            Continue While
        '        ElseIf UseHeader And byteBuf.Count <= 500 Then
        '            '' ヘッダ使用判別後、500バイトまでは無条件で読み込み
        '            Continue While
        '        End If

        '        '' Headerがまだ見つかっていない
        '        If Not UseHeader Then
        '            '' Headerがあるか？
        '            UseHeader = regHead.IsMatch(strBuf.ToString)
        '            If UseHeader Then
        '                '' 見つかった場合は、Headerの終了タグをターゲットに切り替える
        '                regTarget = New System.Text.RegularExpressions.Regex(REG_PAT_HD_ED, RegularExpressions.RegexOptions.Compiled)
        '                Continue While
        '            End If
        '        End If

        '        'If strBuf.ToString.Contains(PROCESS_SOAP_HEADER_END) OrElse strBuf.ToString.Contains(SOAP_HEADER_END) OrElse strBuf.ToString.Contains(SOAP_BODY_START) Then
        '        'If IsCompletedReadHeader(regex, strBuf) Then
        '        If regTarget.IsMatch(strBuf.ToString) Then
        '            ' SoapHeader部分を読込完了
        '            If strBuf.ToString.Contains(IS_COMPRESSION_TRUE) Then
        '                ' "SoapHeaderに、IsCompression="true"を含む場合
        '                ' 圧縮する
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