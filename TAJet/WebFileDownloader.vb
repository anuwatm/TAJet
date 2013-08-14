Imports System.Net
Imports System.IO
Imports System.Net.Mail
Namespace WebTools
    Public Class WebFileDownloader
        Public Event AmountDownloadedChanged(ByVal iNewProgress As Long)
        Public Event FileDownloadSizeObtained(ByVal iFileSize As Long)
        Public Event FileDownloadComplete()
        Public Event FileDownloadFailed(ByVal ex As Exception)
        Private mCurrentFile As String = String.Empty
        Public ReadOnly Property CurrentFile() As String
            Get
                Return mCurrentFile
            End Get
        End Property
        Public Function DownloadFile(ByVal URL As String, ByVal Location As String) As Boolean
            Try
                mCurrentFile = GetFileName(URL)
                Dim WC As New WebClient
                WC.DownloadFile(URL, Location)
                RaiseEvent FileDownloadComplete()
                Return True
            Catch ex As Exception
                RaiseEvent FileDownloadFailed(ex)
                Return False
            End Try
        End Function
        Private Function GetFileName(ByVal URL As String) As String
            Try
                Return URL.Substring(URL.LastIndexOf("/") + 1)
            Catch ex As Exception
                Return URL
            End Try
        End Function
        Public Function DownloadFileWithProgress(ByVal URL As String, ByVal Location As String) As Boolean
            Dim FS As FileStream
            Try
                mCurrentFile = GetFileName(URL)
                Dim wRemote As WebRequest
                Dim bBuffer As Byte()
                ReDim bBuffer(256)
                Dim iBytesRead As Integer
                Dim iTotalBytesRead As Integer
                FS = New FileStream(Location, FileMode.Create, FileAccess.Write)
                wRemote = WebRequest.Create(URL)
                Dim myWebResponse As WebResponse = wRemote.GetResponse
                RaiseEvent FileDownloadSizeObtained(myWebResponse.ContentLength)
                Dim sChunks As Stream = myWebResponse.GetResponseStream
                Do
                    iBytesRead = sChunks.Read(bBuffer, 0, 256)
                    FS.Write(bBuffer, 0, iBytesRead)
                    iTotalBytesRead += iBytesRead
                    If myWebResponse.ContentLength < iTotalBytesRead Then
                        RaiseEvent AmountDownloadedChanged(myWebResponse.ContentLength)
                    Else
                        RaiseEvent AmountDownloadedChanged(iTotalBytesRead)
                    End If
                Loop While Not iBytesRead = 0
                sChunks.Close()
                FS.Close()
                RaiseEvent FileDownloadComplete()
                Return True
            Catch ex As Exception
                If Not (FS Is Nothing) Then
                    FS.Close()
                    FS = Nothing
                End If
                RaiseEvent FileDownloadFailed(ex)
                Return False
            End Try
        End Function
        Public Shared Function FormatFileSize(ByVal Size As Long) As String
            Try
                Dim KB As Integer = 1024
                Dim MB As Integer = KB * KB
                If Size < KB Then
                    Return (Size.ToString("D") & " bytes")
                Else
                    Select Case Size / KB
                        Case Is < 1000
                            Return (Size / KB).ToString("N") & "KB"
                        Case Is < 1000000
                            Return (Size / MB).ToString("N") & "MB"
                        Case Is < 10000000
                            Return (Size / MB / KB).ToString("N") & "GB"
                        Case Else
                            Return Size
                    End Select
                End If
            Catch ex As Exception
                Return Size.ToString
            End Try
        End Function
    End Class

    Public Class webUtilities
        Public Function GetFileNameFromURL(ByVal URL As String) As String
            Try
                Return URL.Substring(URL.LastIndexOf("/") + 1)
            Catch ex As Exception
                Return URL
            End Try
        End Function
        Public Function getWebPage(ByVal URL As String, Optional ByVal streamEndcoding As String = "UTF-8") As String
            Dim strResult As String = ""
            Dim myWebRequest As WebRequest = WebRequest.Create(URL)
            Dim myWebResponse As WebResponse = myWebRequest.GetResponse()

            Dim ReceiveStream As Stream = myWebResponse.GetResponseStream()

            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding(streamEndcoding)

            Dim readStream As New StreamReader(ReceiveStream, encode)
            strResult = "Response stream received"
            Dim read(256) As [Char]
            Dim count As Integer = readStream.Read(read, 0, 256)
            While count > 0
                Dim str As New [String](read, 0, count)
                strResult = strResult & str
                count = readStream.Read(read, 0, 256)
            End While
            readStream.Close()
            Return strResult
        End Function
        Public Function URLDecode(ByVal Source As String) As String
            Dim x As Integer = 0
            Dim CharVal As Byte = 0
            Dim sb As New System.Text.StringBuilder()

            For x = 0 To (Source.Length - 1)
                Dim c As Char = Source(x)
                'Check for space
                If (c = "+") Then
                    sb.Append(" ")
                ElseIf c <> "%" Then
                    'Not hex value so add the chars to string builder.
                    sb.Append(c)
                Else
                    'Convert hex value to char value.
                    CharVal = Int("&H" & Source(x + 1) & Source(x + 2))
                    'Add the chars to string builder.
                    sb.Append(Chr(CharVal))
                    'INC by 2
                    x += 2
                End If
            Next

            'Return the string.
            Return sb.ToString()

        End Function

        Public Function URLEncode(ByVal Source As String) As String
            Dim chars() As Char = Source.ToCharArray()
            Dim sb As New System.Text.StringBuilder()

            For Each c As Char In chars
                'Check for safe chars
                If c Like "[A-Z-a-z-0-9]" Then
                    sb.Append(c)
                ElseIf c = " " Then
                    'Append space char
                    sb.Append("+")
                Else
                    'Get hex value from char
                    Dim sHex As String = Hex(Asc(c))
                    'Pad out left 2 places.
                    sHex = "%" & sHex.PadLeft(2, "0")
                    sb.Append(sHex)
                End If
            Next

            'Clean up
            Erase chars

            'Return string
            Return sb.ToString()

        End Function
    End Class

    Public Class MailUtilities
        Structure mMessage
            Dim subject As String
            Dim body As String
        End Structure
        Structure STMPDesc
            Dim Domain As String
            Dim username As String
            Dim Password As String
            Dim Port As Integer
        End Structure

        Function sendMail(ByVal mailServer As STMPDesc, ByVal mailFrom As String, ByVal mailTo As String, ByVal mMail As mMessage, Optional ByVal mEncoding As String = "UTF-8") As Boolean
            Dim oMail As MailMessage = New MailMessage
            Dim oMailAddress As MailAddress = New MailAddress(mailFrom)

            With oMail
                .To.Add(New MailAddress(mailTo))
                .From = New MailAddress(mailFrom)
                .BodyEncoding = System.Text.Encoding.GetEncoding(mEncoding)
                .Subject = mMail.subject
                .Body = mMail.body
            End With

            Dim oServer As SmtpClient = New SmtpClient(mailServer.Domain, mailServer.Port)
            Dim SMTPUserInfo As New NetworkCredential(mailServer.username, mailServer.Password, mailServer.Domain)
            oServer.UseDefaultCredentials = False
            oServer.Credentials = SMTPUserInfo
            oServer.Send(oMail)
        End Function
    End Class
End Namespace



