Public Class SourceHelper

    Protected genStrategy As IGenSource

    Public Sub New(genStrategy As IGenSource)
        Me.genStrategy = genStrategy
    End Sub

    Public Sub GenegateSource()
        genStrategy.Generate()
    End Sub

End Class
