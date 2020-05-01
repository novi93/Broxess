Imports System.Data

Public Class DataEntity
    Implements ICloneable

    Public MvcData As DataSet
    Public ProxyData As DataSet
    Public ProcessData As DataSet

    Public Function Clone() As Object Implements ICloneable.Clone
        Dim shadownObject As New DataEntity
        shadownObject.MvcData = Me.MvcData.Copy
        shadownObject.ProxyData = Me.ProxyData.Copy
        shadownObject.ProcessData = Me.ProcessData.Copy
        Return shadownObject
    End Function
End Class
