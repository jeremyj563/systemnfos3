Public Class NuGetSearchResult
    Public Property Context As Context
    Public Property Data As Datum()
    Public Property LastReopen As DateTime
    Public Property Index As String
    Public Property TotalHits As String
End Class

Public Class Context
    Public Property Vocab As String
End Class

Public Class Datum
    Public Property ID As String
    Public Property Type As String
    Public Property Authors As Object()
    Public Property Description As String
    Public Property Summary As String
    Public Property IconURL As String
    Public Property Registration As String
    Public Property Tags As Object()
    Public Property Version As String
    Public Property Versions As Version()
End Class

Public Class Version
    Public Property ID As String
    Public Property Downloads As Integer
    Public Property Version As String
    Public Property ContentID As String
End Class