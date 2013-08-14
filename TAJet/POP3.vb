
'***************************************************************************************************************************
'Author: Luke Niland
'Language: vb.net
'Version: 1.0
'Description: Class that connects to the passed mail server, get the number of messages out and the text of the messages
'See RFC 1939 for details of the pop3 specification
'****************************************************************************************************************************
Public Class POP3

    'all the vars for use in the classs
    Dim TCP As Net.Sockets.TcpClient
    Dim POP3Stream As System.IO.Stream
    Dim inStream As System.IO.StreamReader
    Dim strDataIn, strNumMains(2) As String
    Dim intNoEmails As Integer

    'Class to connect to the passed mail server on port 110
    Sub POPConnect(ByVal strServer As String, ByVal strUserName As String, ByVal strPassword As String)

        'connect to the pop3 server over port 110
        Try
            TCP = New Net.Sockets.TcpClient
            TCP.Connect(strServer, 110)

            'create stream into the ip
            POP3Stream = TCP.GetStream
            inStream = New System.IO.StreamReader(POP3Stream)

            'Make sure we get the ok back from the server
            If WaitFor("+OK") = False Then POPErrors("Unexpected Response from mail server when connecting" & vbNewLine & strDataIn)

            'send the email down 
            SendData("USER " & strUserName)
            If WaitFor("+OK") = False Then POPErrors("Unexpected Response from mail server when sending user" & vbNewLine & strDataIn)

            SendData("PASS " & strPassword)
            If WaitFor("+OK") = False Then POPErrors("Unexpected Response from mail server when sending Password" & vbNewLine & strDataIn)

        Catch ex As Exception
            MsgBox(ex.Message)

            Exit Sub
        End Try

    End Sub

    'Function to get the number of mail messages waiting on the server
    Function GetMailStat() As Integer

        'send the stat command and make sure it returns as expected
        Try
            SendData("STAT")
            If WaitFor("+OK") = False Then
                POPErrors("Unexpected Response from mail server when Getting No of Messages" & vbNewLine & strDataIn)
                Exit Function
            Else
                'split up the response. It should be +OK Num of emails size of emails
                strNumMains = Split(strDataIn, " ")
                GetMailStat = strNumMains(1)
                intNoEmails = strNumMains(1)
            End If
        Catch ex As Exception
            MsgBox("Unexpected Error when getting number of emails:" & vbNewLine & ex.Message)
            GetMailStat = 0
            intNoEmails = 0
        End Try

    End Function

    'Output the errors to the user
    Public Function POPErrors(ByVal strMsg As String)

        MsgBox("POP3 ERROR - " & vbNewLine & strMsg, MsgBoxStyle.Critical)

    End Function

    'function to take in what we expect and compare to what we actually get back
    Public Function WaitFor(ByVal strTarget As String) As Boolean

        strDataIn = inStream.ReadLine
        If InStr(strDataIn, strTarget) Then
            WaitFor = True
        Else
            WaitFor = False
        End If

    End Function

    'This function will get the email message of the pop3 server, based on the message number passed
    Public Function GetMailMessage(ByVal intNum As Integer) As String

        Dim strTemp As String
        Dim strEmailMess As String = ""
        Try
            'send the command to the server to return that email back. Command is RETR and the email no ie RETR 1
            SendData("RETR " & Str(intNum))
            'make sure we get a proper response back
            If WaitFor("+OK") = False Then
                POPErrors("Unexpected Response from mail server getting email body" & vbNewLine & strDataIn)
                GetMailMessage = ("No Email was Retrived")
                Exit Function
            End If

            'Should be ok at this point to read in the tcp stream. We read it in until the end of the email
            'whitch is signified by a line containing only a fullpoint(chr46)
            strTemp = inStream.ReadLine

            While (strTemp <> ".")
                strEmailMess = strEmailMess & strTemp & vbCrLf
                strTemp = inStream.ReadLine
            End While

            GetMailMessage = strEmailMess

        Catch ex As Exception
            'just return an error message if we fell over
            GetMailMessage = "No Email was Retrived"
        End Try

    End Function

    'function that will mark an email for deletion. Delete does not occur until the QUIT is issued to the server
    Function MarkForDelete(ByVal intMailItem As Integer)

        Try
            SendData("DELE " & Str(intMailItem))
            If WaitFor("+OK") = False Then POPErrors("Unexpected Response Marking email for deletion" & vbNewLine & strDataIn)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    'Function that will quit the connection to the server (deleting marked mail) and close open readers etc
    Function CloseConn()

        Try
            SendData("QUIT")
            inStream.Close()
            POP3Stream.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    'function to send data down the tcp pipe
    Public Function SendData(ByVal strCommand As String)

        Dim outBuff As Byte()

        outBuff = ConvertStringToByteArray(strCommand & vbCrLf)
        POP3Stream.Write(outBuff, 0, strCommand.Length + 2)

    End Function

    Public Shared Function ConvertStringToByteArray(ByVal stringToConvert As String) As Byte()
        Return (New System.Text.ASCIIEncoding).GetBytes(stringToConvert)
    End Function

End Class

'Class to extract the elements from raw email text
Public Class EmailMessage

    Private m_MessageSource As String

    'function that will call the main proc with what to bring back for everything but the body text
    Public Function ParseEmail(ByVal strMessage As String, ByVal strType As String) As String

        m_MessageSource = strMessage

        'call the parse routine with the pass filed we want
        ParseEmail = ParseHeader(strType)

    End Function

    'Function to parse each of the header parts out of the email
    Private Function ParseHeader(ByVal strHeader As String) As String

        Dim intLenToStart As Integer
        Dim intLenToLineEnd As Integer
        Dim strTmp As String

        intLenToStart = (InStr(m_MessageSource, strHeader) - 1)
        intLenToLineEnd = InStr(Mid(m_MessageSource, intLenToStart), vbCrLf)
        strTmp = m_MessageSource.Substring(intLenToStart, intLenToLineEnd)

        ParseHeader = Replace(strTmp, vbCrLf, "")

    End Function

    'Funtion to parse out the email body 
    Public Function ParseBody() As String

        'To get the body, everything after the first null line of the message is it (rfc822)
        Dim strTmp As String

        'set the temp var to the message body by getting everything after the null line
        strTmp = m_MessageSource.Substring(m_MessageSource.IndexOf(vbCrLf + vbCrLf))

        'get the encoding of the message out, that way we know if we have to run it through the base64 decode
        'routine or not
        If InStr(m_MessageSource, "Content-Transfer-Encoding: base64") Then
            'call the decode routine
            strTmp = DecodeBase64(strTmp)
        End If

        'if the jobs got html content, remove that from the body
        If InStr(strTmp, "------_=_NextPart_") Then
            strTmp = strTmp.Substring(1, strTmp.IndexOf(vbCrLf & vbCrLf & vbCrLf & "------_=_NextPart_"))
        End If


        'Strip out the odd hex that apears at the start and the end
        strTmp = Replace(strTmp, Chr(10) & Chr(9), "")
        strTmp = Replace(strTmp, Chr(13), "")

        ParseBody = Trim(strTmp)

    End Function

    'Function that will decode base64 encoded email body
    Private Function DecodeBase64(ByVal strBody As String)

        Try
            Dim encoding As New System.Text.UTF8Encoding
            Dim Buffer As Byte() = Convert.FromBase64String(strBody)
            DecodeBase64 = encoding.GetString(Buffer)
            encoding = Nothing
            Buffer = Nothing
        Catch ex As Exception
            MsgBox("A problem occured while decoding a base64 email", MsgBoxStyle.Critical)
            DecodeBase64 = "ERROR"
            Exit Function
        End Try

    End Function


End Class