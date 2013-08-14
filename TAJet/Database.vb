Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text
Imports System.IO.FileInfo


Public Class Database
#Region "Dim Variable"
    Dim StringCon As String
    Dim strConfigname As String
    Dim strServername As String
    Dim strUsername As String
    Dim strPassword As String
    Dim strDatabasename As String
#End Region

#Region "Custom properties"
    Public Property ConString() As String
        Get
            Return StringCon
        End Get
        Set(ByVal value As String)
            StringCon = value
        End Set
    End Property
    Public Property ConfigName() As String
        Get
            Return strConfigname
        End Get
        Set(ByVal value As String)
            strConfigname = value
        End Set
    End Property
    Public Property Servername() As String
        Get
            Return strServername
        End Get
        Set(ByVal value As String)
            strServername = value
        End Set
    End Property
    Public Property Username() As String
        Get
            Return strUsername
        End Get
        Set(ByVal value As String)
            strUsername = value
        End Set
    End Property
    Public Property Databasename() As String
        Get
            Return strDatabasename
        End Get
        Set(ByVal value As String)
            strDatabasename = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return strPassword
        End Get
        Set(ByVal value As String)
            strPassword = value
        End Set
    End Property
    Public ReadOnly Property ConnectionSQLServer()
        Get
            Dim strCn As String = "Data Source=" & strServername & ";Initial Catalog=" & strDatabasename & ";User ID=" & strUsername & ";Password=" & strPassword
            Return strCn
        End Get
    End Property
    Public ReadOnly Property ConnectionOleDB()
        Get
            Dim strCn As String = ""
            Return strCn
        End Get
    End Property
#End Region
    Public Sub New()
        StringCon = ""
    End Sub
    
    Function GetConfigValue(ByVal Name As String) As String
        Dim oConfigReader As AppSettingsReader = New AppSettingsReader
        Try
            Dim strValue As String = oConfigReader.GetValue(Name, GetType(System.String))
            Return strValue
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' New object and get Connection String in app.config
    ''' </summary>
    ''' <param name="ConfigName">Name in Application Tag in app.config </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ConfigName As String)
        strConfigname = ConfigName
        StringCon = GetConfigValue(ConfigName)
    End Sub
    ''' <summary>
    ''' New Object and Create Connection String (SQL Server 7)
    ''' </summary>
    ''' <param name="Servername">Server Name (Server Name or IP Address)</param>
    ''' <param name="Username">Username</param>
    ''' <param name="Password">Password</param>
    ''' <param name="Databasename">Database name</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Servername As String, ByVal Username As String, ByVal Password As String, ByVal Databasename As String)
        StringCon = "Data Source=" & Servername & ";Initial Catalog=" & Databasename & ";User ID=" & Username & ";Password=" & Password
        strServername = Servername
        strUsername = Username
        strPassword = Password
        strDatabasename = Databasename
    End Sub
#Region "Function For OleDB provider"
    ''' <summary>
    ''' Create OleDb Connection Object
    ''' </summary>
    ''' <param name="Con">Connection String (String)</param>
    ''' <returns>Object OleDbConnection</returns>
    ''' <remarks></remarks>
    Public Function CreateConnectionOleDB(ByVal Con As String) As OleDbConnection
        Dim Cn As New OleDbConnection(Con)
        Try
            Cn.Open()
            Return Cn
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
    ''' <summary>
    ''' Create OleDb Connection (get Connection String in app.config - ConnectionString Tag)
    ''' </summary>
    ''' <returns>Object OleDbConnection</returns>
    ''' <remarks></remarks>
    Public Function CreateConnectionOleDB() As OleDbConnection
        Dim Cn As OleDbConnection
        If StringCon <> "" Then
            Cn = New OleDbConnection(StringCon)
        ElseIf strConfigname <> "" Then
            'Cn = New OleDbConnection(ConfigurationManager.ConnectionStrings(strConfigname).ConnectionString)
            Cn = New OleDbConnection(GetConfigValue(strConfigname))
        Else
            'Cn = New OleDbConnection(ConfigurationManager.ConnectionStrings("default").ConnectionString)
            Cn = New OleDbConnection(GetConfigValue("default"))
        End If
        Cn.Open()
        Return Cn
    End Function
    ''' <summary>
    ''' Read Data into DataReader
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement (String)</param>
    ''' <param name="Con">OleDbConnection</param>
    ''' <returns>OleDbDataReader</returns>
    ''' <remarks></remarks>
    Public Function DataReaderOleDB(ByVal SQLstring As String, ByVal Con As OleDbConnection) As OleDbDataReader
        Dim Dc As New OleDbCommand(SQLstring, Con)
        Dim dr As OleDbDataReader
        dr = Dc.ExecuteReader()
        Return dr
        Con.Close()
    End Function
    ''' <summary>
    ''' Read Data in DataTable
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement (String)</param>
    ''' <param name="Con">OleDB Connection</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function ReadDataOleDB(ByVal SQLstring As String, ByVal Con As OleDbConnection) As DataTable
        Dim DA As New OleDbDataAdapter(SQLstring, Con)
        Dim DS As New DataSet
        Try
            DA.Fill(DS, "t")
            Return DS.Tables("t")
            DS.Clear()
        Catch ex As Exception
            Return Nothing
        Finally
            Con.Close()
        End Try
    End Function
    Public Function ReadDataOleDB(ByVal SQLstring As String, ByVal StringConnection As String) As DataTable
        Return ReadDataOleDB(SQLstring, CreateConnectionOleDB(StringConnection))
    End Function
    ''' <summary>
    ''' Count Record in SQL Statement
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement</param>
    ''' <param name="Con">OleDbConnection</param>
    ''' <returns>Integer (Support -1 to 2,147,483,647 records)</returns>
    ''' <remarks></remarks>
    Public Function CheckRecordCountOleDB(ByVal SQLstring As String, ByVal Con As OleDbConnection) As Integer
        Dim DA As New OleDbDataAdapter(SQLstring, Con)
        Dim DS As New DataSet
        Dim intCount As Integer

        DA.Fill(DS, "Counter")
        intCount = DS.Tables("Counter").Rows.Count
        Return intCount
    End Function
    ''' <summary>
    ''' Execute SQL Statement
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement</param>
    ''' <param name="Con">OleDbConnection</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>True is Complete or False has any error</remarks>
    Public Function ExecuteSQLCommandOleDB(ByVal SQLstring As String, ByVal Con As OleDbConnection) As Boolean
        Dim dc As New OleDbCommand(SQLstring, Con)
        Try
            dc.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        Finally
            Con.Close()
        End Try
    End Function
    Public Function ExecuteSQLCommandOleDB(ByVal SQLstring As String, ByVal StringConnection As String) As Boolean
        Return ExecuteSQLCommandOleDB(SQLstring, CreateConnectionOleDB(StringConnection))
    End Function
#End Region
#Region "Function For SQLServer Provider"
    ''' <summary>
    ''' Create Connection String
    ''' </summary>
    ''' <param name="Servername">Server Name (Support Name or IP Address)</param>
    ''' <param name="Username">Username</param>
    ''' <param name="Password">Password</param>
    ''' <param name="Databasename">Database Name</param>
    ''' <returns>Connection String (String)</returns>
    ''' <remarks></remarks>
    Public Function CreateConnnectionString(ByVal Servername As String, ByVal Username As String, ByVal Password As String, ByVal Databasename As String) As String
        Return "Data Source=" & Servername & ";Initial Catalog=" & Databasename & ";User ID=" & Username & ";Password=" & Password
    End Function
    ''' <summary>
    ''' Create SQL Connection
    ''' </summary>
    ''' <param name="Con">Connection String (String)</param>
    ''' <returns>SQL Connection</returns>
    ''' <remarks></remarks>
    Public Function CreateConnectionDB(ByVal Con As String) As SqlConnection
        Dim Cn As New SqlConnection(Con)

        Try
            Cn.Open()
            Return Cn
        Catch ex As Exception
            Return Nothing
        Finally
            'Cn.Close()
        End Try
    End Function
    ''' <summary>
    ''' Create SQL Connection
    ''' </summary>
    ''' <returns>SQL Connection</returns>
    ''' <remarks>Get ConnectionString in app.config (Default Value)</remarks>
    Public Function CreateConnectionDB() As SqlConnection
        Dim Cn As SqlConnection
        If StringCon <> "" Then
            Cn = New SqlConnection(StringCon)
        ElseIf strConfigname <> "" Then
            Cn = New SqlConnection(GetConfigValue(strConfigname))
        Else
            Cn = New SqlConnection(GetConfigValue("default"))
        End If

        Try
            Cn.Open()
            Return Cn
        Catch ex As Exception
            Return Nothing
        Finally
            Cn.Close()
        End Try
    End Function
    ''' <summary>
    ''' Read Data get into SQLDataReader
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement</param>
    ''' <param name="Con">Sql Connection</param>
    ''' <returns>SQL Data reader</returns>
    ''' <remarks></remarks>
    Public Function DataReader(ByVal SQLstring As String, ByVal Con As SqlConnection) As SqlDataReader
        Dim Dc As New SqlCommand(SQLstring, Con)
        Dim dr As SqlDataReader

        Try
            dr = Dc.ExecuteReader()
            Return dr
        Catch ex As Exception
            Return Nothing
        Finally
            dr.Close()
            Con.Close()
        End Try

    End Function
    ''' <summary>
    ''' Read SQL Statement and get into DataTable
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement (String)</param>
    ''' <param name="Con">Sql Connection</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function ReadData(ByVal SQLstring As String, ByVal Con As SqlConnection) As DataTable
        Dim DA As New SqlDataAdapter(SQLstring, Con)
        Dim DS As New DataSet
        Try
            DA.Fill(DS, "t")
            Return DS.Tables("t")
            DS.Clear()
        Catch ex As Exception
            Return Nothing
        Finally
            If Not (Con Is Nothing) Then
                Con.Close()
            End If
        End Try
    End Function
    Public Function Readdata(ByVal SQLString As String, ByVal StringConnection As String) As DataTable
        Return Readdata(SQLString, CreateConnectionDB(StringConnection))
    End Function
    ''' <summary>
    ''' Count Record in SQL Statement
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement (String)</param>
    ''' <param name="Con">Sql Connection</param>
    ''' <returns>Integer (Support -1 to 2,147,483,647 records)</returns>
    ''' <remarks></remarks>
    Public Function CheckRecordCount(ByVal SQLstring As String, ByVal Con As SqlConnection) As Integer
        Dim DA As New SqlDataAdapter(SQLstring, Con)
        Dim DS As New DataSet
        Dim intCount As Integer

        Try
            DA.Fill(DS, "Counter")
            intCount = DS.Tables("Counter").Rows.Count
            DS.Clear()
            Return intCount
        Catch ex As Exception
            Return -1
        Finally
            Con.Close()
        End Try
    End Function
    ''' <summary>
    ''' Execute SQL Statement
    ''' </summary>
    ''' <param name="SQLstring">SQL Statement</param>
    ''' <param name="Con">OleDbConnection</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>True is Complete or False has any error</remarks>
    Public Function ExecuteSQLCommand(ByVal SQLstring As String, ByVal Con As SqlConnection) As Integer
        Dim dc As New SqlCommand(SQLstring, Con)
        Dim lngCounter As Integer = 0
        Try
            lngCounter = dc.ExecuteNonQuery()
            Return lngCounter
        Catch ex As Exception
            Return 0
        Finally
            Con.Close()
        End Try
    End Function
    Public Function ExecuteSQLCommand(ByVal SQLstring As String, ByVal StringConnection As String) As Integer
        Return ExecuteSQLCommand(SQLstring, CreateConnectionDB(StringConnection))
    End Function
#End Region
End Class

Public Class ManageFormat
    Enum TypeofDB
        DBAccess = 1
        DBSQLServer = 2
    End Enum
    Enum TypeYear
        yearBudda = 1
        yearCrist = 2
    End Enum
    Enum TypeConvertChar
        UTF_Ascii = 1
        Ascii_UTF = 2
    End Enum
    ''' <summary>
    ''' Format String Value
    ''' </summary>
    ''' <param name="Value">String Value</param>
    ''' <returns>String For SQL Statement</returns>
    ''' <remarks>Format 'String' If Value is '' will return NULL</remarks>
    Public Function FormatValue(ByVal Value As String, Optional ByVal UseDoublequotationmarks As Boolean = False) As String
        If Value = "" Then
            Return "null"
        Else
            If UseDoublequotationmarks Then
                Return """" & Value & """"
            Else
                Return "'" & Value & "'"
            End If
            'Return "'" & Value.Replace("'", "&#146;") & "'"
        End If
    End Function
    ''' <summary>
    ''' Format Number (Type Double)
    ''' </summary>
    ''' <param name="Value">Number value (Type DOUBLE)</param>
    ''' <returns>String For SQL Statement</returns>
    ''' <remarks>Format Number If Value is DBNull will return NULL</remarks>
    Public Function FormatValue(ByVal Value As Double) As String
        If IsDBNull(Value) Then
            Return "null"
        Else
            Return Value
        End If
    End Function
    ''' <summary>
    ''' Format Number (Type Integer)
    ''' </summary>
    ''' <param name="Value">Number value (Type INTEGER)</param>
    ''' <returns>String For SQL Statement</returns>
    ''' <remarks>Format Number If Value is DBNull will return NULL</remarks>
    Public Function FormatValue(ByVal Value As Integer) As String
        If IsDBNull(Value) Then
            Return "null"
        Else
            Return Value
        End If
    End Function
    ''' <summary>
    ''' Format Date
    ''' </summary>
    ''' <param name="Value">Date Value</param>
    ''' <param name="DBType">Type of Database (Access or SQL Server)</param>
    ''' <returns>String For SQL Statement</returns>
    ''' <remarks>If Access format #Date# and SQL Server format 'Date'.If DBNull will return NULL</remarks>
    Public Function Formatvalue(ByVal Value As Date, ByVal DBType As TypeofDB) As String
        Select Case DBType
            Case TypeofDB.DBAccess
                If IsDBNull(Value) Then
                    Return "null"
                Else
                    Return "#" & FormatDate(Value) & "#"
                End If
            Case TypeofDB.DBSQLServer
                If IsDBNull(Value) Then
                    Return "null"
                Else
                    Return "'" & FormatDate(Value) & "'"
                End If
            Case Else
                Return Value
        End Select
    End Function
    ''' <summary>
    ''' Format Date
    ''' </summary>
    ''' <param name="Value">Date Value</param>
    ''' <param name="ConvertTo">Convert to Year (YearBudda or yearCrist)</param>
    ''' <returns>String For SQL Statement</returns>
    ''' <remarks>Set Date Format is Year/Month/Date</remarks>
    Public Function FormatDate(ByVal Value As Date, Optional ByVal ConvertTo As TypeYear = TypeYear.yearBudda) As String
        Dim strTemp As String = ""
        Select Case ConvertTo
            Case TypeYear.yearBudda
                If Value.Year < 2500 Then
                    strTemp = ((Value.Year) + 543 & "/" & Value.Month & "/" & Value.Day).ToString
                Else
                    strTemp = ((Value.Year) & "/" & Value.Month & "/" & Value.Day).ToString
                End If
            Case TypeYear.yearCrist
                If Value.Year > 2500 Then
                    strTemp = ((Value.Year) - 543 & "/" & Value.Month & "/" & Value.Day).ToString
                Else
                    strTemp = ((Value.Year) & "/" & Value.Month & "/" & Value.Day).ToString
                End If
        End Select

        Return strTemp
    End Function
    ''' <summary>
    ''' Convert String Between ASCII and Unicode
    ''' </summary>
    ''' <param name="StringSource">String Source (String)</param>
    ''' <param name="ConvertType">Type Convert (Ascii or Unicode)</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function ConvertStringCode(ByVal StringSource As String, Optional ByVal ConvertType As TypeConvertChar = TypeConvertChar.UTF_Ascii) As String
        Dim enAscii As Encoding = Encoding.GetEncoding("windows-874")
        Dim enUnicode As Encoding = Encoding.UTF8
        Select Case ConvertType
            Case TypeConvertChar.UTF_Ascii
                Dim unicodeBytes As Byte() = enUnicode.GetBytes(StringSource)

                Dim asciiBytes As Byte() = Encoding.Convert(enUnicode, enAscii, unicodeBytes)

                Dim asciiChars(enAscii.GetCharCount(asciiBytes, 0, asciiBytes.Length)) As Char
                enAscii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
                Dim asciiString As New String(asciiChars)

                Return asciiString
            Case TypeConvertChar.Ascii_UTF
                Dim asciiBytes As Byte() = enAscii.GetBytes(StringSource)
                Dim unicodeBytes As Byte() = Encoding.Convert(enAscii, enUnicode, asciiBytes)
                Dim UTFChars(enUnicode.GetCharCount(unicodeBytes, 0, unicodeBytes.Length)) As Char
                enUnicode.GetChars(unicodeBytes, 0, unicodeBytes.Length, UTFChars, 0)
                Dim UTFString As New String(UTFChars)

                Return UTFString
        End Select
    End Function
End Class

Public Class Utilities
    Enum TypeCheck
        CheckOnly = 1
        AutoCreate = 2
        AutoDelete = 3
    End Enum
    ''' <summary>
    ''' Check Folder Exists
    ''' </summary>
    ''' <param name="Path">Path folder want to check</param>
    ''' <param name="CheckState">State check (Check only ,Automatic create or Automatic Delete)</param>
    ''' <returns>Boolean (True is Complete or False is any error)</returns>
    ''' <remarks>If Folder is not Exist and TypeCheck is AutoCreate it will create folder OR TypeCheck is AutoDelete it will delete folder</remarks>
    Public Function CheckFolderExists(ByVal Path As String, Optional ByVal CheckState As TypeCheck = TypeCheck.CheckOnly) As Boolean
        Dim blnStatus As Boolean = False
        Dim oFolder As DirectoryInfo = New DirectoryInfo(Path)

        Select Case CheckState
            Case TypeCheck.CheckOnly
                If oFolder.Exists() Then
                    blnStatus = True
                End If
            Case TypeCheck.AutoCreate
                If oFolder.Exists() Then
                    blnStatus = True
                Else
                    Try
                        oFolder.Create()
                        blnStatus = True
                    Catch ex As Exception
                        blnStatus = False
                    End Try
                End If
            Case TypeCheck.AutoDelete
                If oFolder.Exists Then
                    Try
                        oFolder.Delete()
                        blnStatus = True
                    Catch ex As Exception
                        blnStatus = False
                    End Try
                End If
        End Select
        Return blnStatus
    End Function
    ''' <summary>
    ''' Get Value in app.config 
    ''' </summary>
    ''' <param name="ConfigName">Config Name (name in Application Setting)</param>
    ''' <returns>Value is String</returns>
    ''' <remarks></remarks>
    Public Function getConfigValue(ByVal ConfigName As String) As String
        Dim oConfigReader As AppSettingsReader = New AppSettingsReader
        Return oConfigReader.GetValue(ConfigName, GetType(System.String))
    End Function
    ''' <summary>
    ''' Update value in app.config
    ''' </summary>
    ''' <param name="configName">Config Name (name in Application Setting)</param>
    ''' <param name="valueConfig">New Value</param>
    ''' <param name="configPath">Tag Config (Default is appSettings)</param>
    ''' <returns>Boolean (TRUE is Complete or FALSE is any ERROR)</returns>
    ''' <remarks></remarks>
    Public Function setConfigValue(ByVal configName As String, ByVal valueConfig As String, Optional ByVal configPath As String = "appSettings") As Boolean
        Dim blnStatus As Boolean = False
        Dim fileConfig As String = My.Application.Info.DirectoryPath() & "\" & My.Application.Info.AssemblyName & ".config"
        Dim oXml As Xml.XmlDocument = New Xml.XmlDocument()

        Try
            oXml.Load(fileConfig)
            For Each oElement As Xml.XmlElement In oXml.DocumentElement
                If oElement.Name = configPath Then
                    For Each oNode As Xml.XmlNode In oElement.ChildNodes
                        If oNode.Attributes(0).Value = configName Then
                            oNode.Attributes(1).Value = valueConfig
                            Exit For
                        End If
                    Next
                End If
            Next
            oXml.Save(fileConfig)
            blnStatus = True
        Catch ex As Exception
            blnStatus = False
        End Try

        Return blnStatus
    End Function
    ''' <summary>
    ''' Create new Folder
    ''' </summary>
    ''' <param name="Path">Path New Folder is want to Create</param>
    ''' <param name="subDirectory">subDirectory  is want to Create (Optional)</param>
    ''' <returns>Boolean (TRUE is Create SUCCESS or FALSE is any Error)</returns>
    ''' <remarks></remarks>
    Public Function CreateFolder(ByVal Path As String, Optional ByVal subDirectory As String = "") As Boolean
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
    'Public Function ShowMessage(ByVal Message As String, ByVal Caption As String, Optional ByVal ButtonMessage As MessageBoxButtons = MessageBoxButtons.OK, Optional ByVal IconMessasge As MessageBoxIcon = MessageBoxIcon.Information) As DialogResult
    '    Return MessageBox.Show(Message, Caption, ButtonMessage, IconMessasge)
    'End Function

    ''' <summary>
    ''' Create New Text File
    ''' </summary>
    ''' <param name="textString">Value in text File</param>
    ''' <param name="PathFile">Path File to save</param>
    ''' <returns>Boolean (TRUE is create SUCCESS or False is any ERROR)</returns>
    ''' <remarks>Text File create is Unicode</remarks>
    Public Function CreateTextFile(ByVal textString As String, ByVal PathFile As String, Optional ByVal Encoding As String = "UTF-8") As Boolean
        Try
            Dim objWriter As New StreamWriter(PathFile, False, System.Text.Encoding.GetEncoding(Encoding))
            objWriter.Write(textString)
            objWriter.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function createXMLFile(ByVal XMLString As String, ByVal Pathfile As String, Optional ByVal Encoding As String = "UTF-8") As String
        Try
            Dim objWriter As New StreamWriter(Pathfile, False, System.Text.Encoding.GetEncoding(Encoding))
            objWriter.Write("<?xml version=""1.0"" encoding=""" & Encoding & """ ?>" & vbCrLf & "<data>" & vbCrLf & XMLString & vbCrLf & "</data>")
            objWriter.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function deCrypto(ByVal AlgoType As Crypto.Algorithm, ByVal CryptoType As Crypto.EncodingType, ByVal Source As String, ByVal Key As String) As String
        Crypto.Key = Key
        Crypto.EncryptionAlgorithm = AlgoType
        Crypto.Encoding = CryptoType
        Crypto.Content = Source
        If Crypto.DecryptString Then
            Return Crypto.Content
        Else
            Return Crypto.CryptoException.Message
        End If
    End Function
    Public Function checkFileExists(ByVal fullFilename As String) As Boolean
        Dim oFileInfo As FileInfo = New FileInfo(fullFilename)
        If oFileInfo.Exists() Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function OpenTextFile(ByVal Filename As String, Optional ByVal Encode As String = "UTF-8") As String
        Dim strText As String = ""
        Dim strPath As String = Filename
        Dim oUtilities As Utilities = New Utilities

        Try
            Dim oStreamReader As New StreamReader(strPath, Encoding.GetEncoding(Encode))
            With oStreamReader
                Do While .Peek() >= 0
                    strText = strText & .ReadLine() 'oFormat.ConvertStringCode(.ReadLine(), ManageFormat.TypeConvertChar.Ascii_UTF)
                Loop
                oStreamReader.Close()
            End With
        Catch ex As Exception
            strText = ""
        End Try
        Return strText
    End Function
    Public Function getProvideinConfig(ByVal ServerConfig As String, ByVal UsernameConfig As String, ByVal PasswordConfig As String, ByVal DBConfig As String, ByVal HashKey As String, Optional ByVal AlgorithmType As Crypto.Algorithm = Crypto.Algorithm.DES, Optional ByVal EncodeType As Crypto.EncodingType = Crypto.EncodingType.HEX) As String
        Dim hKey As String = getConfigValue(HashKey)
        Dim strServer As String = deCrypto(AlgorithmType, EncodeType, getConfigValue("Server"), hKey)
        Dim strUsername As String = deCrypto(AlgorithmType, EncodeType, getConfigValue("Username"), hKey)
        Dim strPassword As String = deCrypto(AlgorithmType, EncodeType, getConfigValue("Password"), hKey)
        Dim strDBname As String = deCrypto(AlgorithmType, EncodeType, getConfigValue("DBname"), hKey)
        Dim DB As New Database()
        Return DB.CreateConnnectionString(strServer, strUsername, strPassword, strDBname)
    End Function
    Public Function CryptText(ByVal SourceText As String, ByVal HashKey As String, Optional ByVal AlgorithmType As Crypto.Algorithm = Crypto.Algorithm.DES, Optional ByVal EncodeType As Crypto.EncodingType = Crypto.EncodingType.HEX) As String
        Dim strCrypt As String = ""

        Return strCrypt
    End Function
    Public Function DECryptText(ByVal SourceText As String, ByVal HashKey As String, Optional ByVal AlgorithmType As Crypto.Algorithm = Crypto.Algorithm.DES, Optional ByVal EncodeType As Crypto.EncodingType = Crypto.EncodingType.HEX) As String
        Dim strDECrypt As String = ""

        Return strDECrypt
    End Function
    Public Function ConvertSize(ByVal int_64 As Int64)
        Dim KiloBytes As Int64 = Convert.ToInt64(int_64 / 1024)
        Dim MegaBytes As Int64 = Convert.ToInt64(KiloBytes / 1024)
        Dim GigBytes As Int64 = Convert.ToInt64(MegaBytes / 1024)
        Dim TeraBytes As Int64 = Convert.ToInt64(GigBytes / 1024)
        If KiloBytes >= 1024 Then
            If MegaBytes >= 1024 Then
                'GigaBytes
                If GigBytes >= 1024 Then
                    Return Math.Round(TeraBytes) & " TB".ToString
                Else
                    Return Math.Round(GigBytes) & " GB".ToString
                End If
            Else
                'Megabytes
                Return Math.Round(MegaBytes) & " MB".ToString
            End If
        Else
            'Kilobytes
            Return Math.Round(KiloBytes) & " KB".ToString
        End If
    End Function
End Class
