Imports System.IO
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols


Public Class ErrorDetail
    Dim strID As String
    Dim strDesc As String
    Dim strFunctionName As String
    Dim strOther As String
    Dim dtaTimeStamp As DateTime

    Public Sub New()
        strID = ""
        strDesc = ""
        strFunctionName = ""
        strOther = ""
        dtaTimeStamp = Date.Now.ToString
    End Sub
    Public Sub New(ByVal ID As String, ByVal Description As String, Optional ByVal Functionname As String = "", Optional ByVal Other As String = "")
        strID = ID
        strDesc = Description
        strFunctionName = Functionname
        strOther = Other
        dtaTimeStamp = Date.Now.ToString
    End Sub
    Public Property ID() As String
        Get
            Return strID
        End Get
        Set(ByVal value As String)
            strID = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return strDesc
        End Get
        Set(ByVal value As String)
            strDesc = value
        End Set
    End Property

    Public Property Other() As String
        Get
            Return strOther
        End Get
        Set(ByVal value As String)
            strOther = value
        End Set
    End Property

    Public Property FunctionName() As String
        Get
            Return strFunctionName
        End Get
        Set(ByVal value As String)
            strFunctionName = value
        End Set
    End Property

    Public ReadOnly Property TimeStamp() As DateTime
        Get
            Return dtaTimeStamp
        End Get
    End Property
    Private Function CreateFolder(ByVal Path As String, Optional ByVal subDirectory As String = "") As Boolean
        Dim oDir As New DirectoryInfo(Path)
        If oDir.Exists = False Then
            oDir.Create()
        End If
        If subDirectory.Length > 0 Then
            oDir = New DirectoryInfo(Path & "\" & subDirectory)
            If oDir.Exists = False Then
                oDir.Create()
            End If
        End If
        Return True
    End Function
    Public Sub WriteLog()
        Dim oUtilities As Utilities = New Utilities
        Dim strPath As String = oUtilities.getConfigValue("ErrorLog").ToString
        CreateFolder(strPath)
        strPath = strPath & "\" & Date.Now.Year.ToString & "_" & Right("0" & Date.Now.Month.ToString, 2) & ".txt"
        Dim oStreamWriter As New StreamWriter(strPath, True)
        Dim strString As String = strID & " | " & strDesc & " | " & strFunctionName & " | " & dtaTimeStamp
        oStreamWriter.WriteLine(strString)
        oStreamWriter.Close()
    End Sub
    Public Sub WriteLog(ByVal pathFile As String)
        Dim oUtilities As Utilities = New Utilities
        CreateFolder(pathFile)
        pathFile = pathFile & "\" & Date.Now.Year.ToString & "_" & Right("0" & Date.Now.Month.ToString, 2) & ".txt"
        Dim oStreamWriter As New StreamWriter(pathFile, True)
        Dim strString As String = strID & " | " & strDesc & " | " & strFunctionName & " | " & dtaTimeStamp
        oStreamWriter.WriteLine(strString)
        oStreamWriter.Close()
    End Sub
End Class
