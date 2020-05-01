Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Xml.Serialization
Imports System.IO

Public Class TableInfoEntity
    Implements INotifyPropertyChanged

    Public no As Integer
    Private _no As Integer
    Public WriteOnly Property rowNo() As Integer
        Set(ByVal value As Integer)
            _no = value
            OnPropertyChanged((New PropertyChangedEventArgs("rowNo")))
        End Set
    End Property

    Private _tableName As String
    Public Property TableName() As String
        Get
            Return _tableName
        End Get
        Set(ByVal value As String)
            _tableName = value.Trim
        End Set
    End Property

    Private _condition As String
    Public Property Condition() As String
        Get
            Return _condition
        End Get
        Set(ByVal value As String)
            _condition = value.Trim
        End Set
    End Property
    Private _order As String
    Public Property Order() As String
        Get
            Return _order
        End Get
        Set(ByVal value As String)
            _order = value.Trim
        End Set
    End Property

    'Private _NumbOfRecord As String

    '<DisplayName("Count")>
    'Public ReadOnly Property NumbOfRecord() As String
    '    Get
    '        Return _NumbOfRecord
    '    End Get
    '    'Set(ByVal value As String)
    '    '    _NumbOfRecord = value
    '    '    OnPropertyChanged((New PropertyChangedEventArgs("NumbOfRecord")))
    '    'End Set
    'End Property

    'Public Sub SetNumbOfRecord(ByVal value As String)
    '    _NumbOfRecord = value
    '    OnPropertyChanged((New PropertyChangedEventArgs("NumbOfRecord")))
    'End Sub

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(ByVal e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Public Function createQuerry() As String
        If String.IsNullOrWhiteSpace(TableName) Then
            Dim ex As New Exception("TableName is empty")
            ex.Data.Add("index", no)
            Throw ex
        End If

        Return String.Format("SELECT * FROM {0} {1} {2}",
                              _tableName,
                              If(String.IsNullOrWhiteSpace(_condition), "", " WHERE " & _condition),
                              If(String.IsNullOrWhiteSpace(_order), "", " ORDER BY " & _order))
    End Function
End Class

Class TableInfoHelper
    Dim _tableInfoCollection As ObservableCollection(Of TableInfoEntity)

    Dim _conString As String = String.Empty

    Dim _dataPath As String = String.Empty

    Public Sub New()
        _tableInfoCollection = New ObservableCollection(Of TableInfoEntity)()
        AddHandler _tableInfoCollection.CollectionChanged, AddressOf _tableInfoCollection_CollectionChanged
    End Sub

    Public Function getTableInfoList() As ObservableCollection(Of TableInfoEntity)
        Return _tableInfoCollection
    End Function


    Public Function getConnectionString() As String
        Return _conString

    End Function

    Public Function getDataPath() As String
        Return _dataPath

    End Function

    Public Sub setConnectionString(val As String)
        If val Is Nothing OrElse String.IsNullOrEmpty(val) Then
            val = ""
        End If
        _conString = val

    End Sub

    Public Sub setDataPath(val As String)
        If val Is Nothing OrElse String.IsNullOrEmpty(val) Then
            val = ""
        End If
        _dataPath = val
    End Sub

    Private Sub _tableInfoCollection_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs)
        For i = 0 To _tableInfoCollection.Count - 1
            _tableInfoCollection(i).rowNo = (i + 1)
        Next
    End Sub

    Public Function isEmptyOrWhiteSpace() As Boolean
        Return Me.Count() = 0

    End Function
    Public Function Count() As Integer
        Return _tableInfoCollection.Where(Function(x)
                                              Return Not String.IsNullOrWhiteSpace(x.TableName)
                                          End Function).
                                      Count()
    End Function
    Public Function LoadData(repo As Repository)

        Dim ds As New DataSet
        For Each tInfo In _tableInfoCollection
            If String.IsNullOrWhiteSpace(tInfo.TableName) Then
                Continue For
            End If
            Dim dt = repo.GetData(tInfo.createQuerry()).Copy
            dt.TableName = tInfo.TableName
            ds.Tables.Add(dt)
            'tInfo.SetNumbOfRecord(dt.Rows.Count)
        Next
        Return ds
    End Function

    Public Sub SaveSetting()
        Dim xs As New XmlSerializer(GetType(ObservableCollection(Of TableInfoEntity)))
        Using wr As New StreamWriter("frmExport_Grid.xml")
            xs.Serialize(wr, _tableInfoCollection)
        End Using
    End Sub
    Public Sub LoadSetting()
        Dim xs As New XmlSerializer(GetType(ObservableCollection(Of TableInfoEntity)))
        If Not File.Exists("frmExport_Grid.xml") Then
            Return
        End If
        Using rd As New StreamReader("frmExport_Grid.xml")
            _tableInfoCollection = xs.Deserialize(rd)
        End Using
    End Sub
End Class
