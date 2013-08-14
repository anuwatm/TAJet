Imports System.Data
Imports System.Xml
Imports System.Xml.XPath


Public Class RSSfeed
    Structure rssProperties
        Dim Title As String
        Dim Channel As String
        Dim lastUpdate As String
        Dim Language As String
        Dim itemRSS As DataTable
    End Structure
    Function createData() As DataTable
        Dim DT As DataTable = New DataTable
        Dim DC As DataColumn
        With DT
            DC = New DataColumn("Title")
            .Columns.Add(DC)
            DC = New DataColumn("Desc")
            .Columns.Add(DC)
            DC = New DataColumn("Link")
            .Columns.Add(DC)

        End With
    End Function
    Public Function loadRSSS(ByVal URL As String) As rssProperties
        Dim oRSS As rssProperties = New rssProperties
        Dim oXML As XmlDocument = New XmlDocument

        oXML.Load(URL)
        Dim oNavi As XPathNavigator = oXML.CreateNavigator
        Try
            Dim oNodes As XPathNavigator = oNavi.Select("/rss/channel/item/title")

        Catch ex As Exception

        End Try
        Return oRSS
    End Function
End Class
