Imports SiU.Process.Common.Functions.CommonMethod
Imports SiU.Process.Common.Functions.SystemValue
Imports System.Xml.Serialization
Imports System.Collections
Imports System.IO
Imports System.Xml

Public Class DataConfigManager
    Public Const CONFIG_FILE = "_config.xml"
    Public Shared Function GetSampleDataConfig() As DataConfig
        Dim a As New DataConfig
        a.TableName = "Programs"
        a.PGCodeColumn = "Id"
        a.PGCodeSource = "01-00000-02"
        a.PGCodeTargert = "222"
        a.Sort.Add("Id")
        a.Where.Add(a.PGCodeColumn, fSqlStr(a.PGCodeSource))
        Return a
    End Function

    Public Shared Sub SaveSetting(objList As List(Of DataConfig))
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(List(Of DataConfig)))
        Using writer As TextWriter = New StreamWriter(CONFIG_FILE)
            serializer.Serialize(writer, objList)
            writer.Close()
        End Using
    End Sub

    Public Shared Function LoadSetting() As List(Of DataConfig)
        Dim rs As List(Of DataConfig)
        Dim serialiser As XmlSerializer = New XmlSerializer(GetType(List(Of DataConfig)))
        Using reader As TextReader = New StreamReader(CONFIG_FILE)
            rs = DirectCast(serialiser.Deserialize(reader), List(Of DataConfig))
            reader.Close()
        End Using
        Return rs
    End Function
End Class
