Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Xml.Serialization
Imports System.Xml

<XmlRoot("dictionary")>
Public Class SerializableDictionary(Of TKey, TValue)
    Inherits Dictionary(Of TKey, TValue)
    Implements IXmlSerializable

#Region "IXmlSerializable Members"

    Public Function GetSchema() As Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
        Return Nothing
    End Function

    Public Sub ReadXml(reader As Xml.XmlReader) Implements IXmlSerializable.ReadXml
        Dim keySerializer As XmlSerializer = New XmlSerializer(GetType(TKey))
        Dim valueSerializer As XmlSerializer = New XmlSerializer(GetType(TValue))

        Dim wasEmpty As Boolean = reader.IsEmptyElement
        reader.Read()

        If wasEmpty Then
            Return
        End If

        While reader.NodeType <> XmlNodeType.EndElement
            reader.ReadStartElement("item")
            reader.ReadStartElement("key")
            Dim key As TKey = DirectCast(keySerializer.Deserialize(reader), TKey)
            reader.ReadEndElement()
            reader.ReadStartElement("value")
            Dim value As TValue = DirectCast(valueSerializer.Deserialize(reader), TValue)
            reader.ReadEndElement()
            reader.ReadEndElement()

            Me.Add(key, value)

            reader.MoveToContent()
        End While

        If (reader.NodeType = XmlNodeType.EndElement) Then
            reader.ReadEndElement()
        Else
            Throw New XmlException("Error in Deserialization of SerializableDictionary")
        End If
    End Sub

    Public Sub WriteXml(writer As Xml.XmlWriter) Implements IXmlSerializable.WriteXml
        Dim keySerializer As XmlSerializer = New XmlSerializer(GetType(TKey))
        Dim valueSerializer As XmlSerializer = New XmlSerializer(GetType(TValue))

        writer.WriteStartElement("item")
        For Each key As TKey In Me.Keys
            writer.WriteStartElement("key")
            keySerializer.Serialize(writer, key)
            writer.WriteEndElement()

            writer.WriteStartElement("value")
            Dim value As TValue = Me(key)
            valueSerializer.Serialize(writer, value)
            writer.WriteEndElement()
        Next

        writer.WriteEndElement()
    End Sub
#End Region

End Class