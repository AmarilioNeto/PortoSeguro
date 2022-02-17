
Imports System.IO
Imports System.Net
Imports System.Xml
Imports System.Text
Imports System.Net.Http
Imports System.Web
Imports HtmlAgilityPack
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net.WebRequest






Namespace Acesso
    Public Class PortoSeguro


        Public Property Url As String = ""
        Public Property Host As String = ""
        Public Property Lang As String = ""
        Public Property Referer As String = ""
        Public Property ContentType As String = ""
        Public Property Cookie As String = ""
        Public Property Origin As String = ""
        Public Property ContentLength As String = ""
        Public Property CookieName As String = "Set-Cookie"
        Public Property Postdata As String = ""
        Public Property TimeOut As Integer = 0
        Public Property Encoding As Encoding
        Public Property QueryString As String = "?"
        Public Property CorrigirHTML As Boolean = False
        Public Property AutoRedirect As Boolean = True
        Public Property Location As String = ""
        Public Property Authorization As String = ""
        Public Property UpgradeInsecureRequests As String = ""
        Public Property AllowAutoRedirect As Boolean = True
        Public Property MaximumAutomaticRedirections = 0
        Public Property GrauInstancia As String = ""
        Public Property CacheControl As String = ""
        Public Property XMLHttpRequest As String = ""
        Public Property AcceptTransferEncoding As String = ""
        Public Property userAgent As String = ""
        Public Property response = Nothing
        Public Property Accept As String = ""
        Public Property AcceptEncoding As String = ""
        Public Property AcceptLanguage As String = ""
        Public Property Connection As String = ""
        Public Property documentCode401 As Boolean = False
        Public Property captchaToken As String()
        Public Property CookieContainer As CookieContainer
        Public Property ContentTypeResponse As String = ""
        Public Property Captcha As String = ""
        Public Property jsonDataBytes As String = ""
        Public Property [mod] As String = ""
        Public Property comp As String = ""
        Public Property user As String = ""
        Public Property pass As String = ""
        Public Property dump As String = ""
        Public Property AutomaticDecompression As String = ""
        Public Property GetResponse As Object




        Public Function CriarPostHttpRequestAsDocument() As HtmlDocument

            Dim request As HttpWebRequest = Nothing
            Dim response As HttpWebResponse = Nothing

            If QueryString.StartsWith("?") And Not QueryString.Equals("?") Then
                Url &= QueryString.TrimEnd("&")
            End If

            request = HttpWebRequest.Create(Url)
            request.CookieContainer = CookieContainer
            request.Method = "POST"
            request.AllowAutoRedirect = AllowAutoRedirect

            If Not String.IsNullOrEmpty(Cookie) Then
                request.Headers.Add(HttpRequestHeader.Cookie, Cookie)
            End If

            If TimeOut <> 0 Then
                request.Timeout = TimeOut
            End If

            If Not String.IsNullOrEmpty(Host) Then
                request.Host = Host
            End If

            request.Headers.Add(HttpRequestHeader.AcceptLanguage, Lang)

            If UpgradeInsecureRequests <> "" Then
                request.Headers.Add("Upgrade-Insecure-Requests", UpgradeInsecureRequests)
            End If

            If Captcha <> "" Then
                request.Headers.Add("captcha", Captcha)
            End If

            If CacheControl <> "" Then
                request.Headers.Add("Cache-Control", CacheControl)
            End If
            If AcceptTransferEncoding <> "" Then
                request.Headers.Add("Accept-Transfer-Encoding", AcceptTransferEncoding)
            End If



            If AcceptEncoding <> "" Then
                request.Headers.Add("Accept-Encoding", AcceptEncoding)
            End If

            request.Referer = Referer

            If userAgent <> "" Then
                request.UserAgent = userAgent
            Else
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; rv:26.0) Gecko/20100101 Firefox/26.0"
            End If
            request.ContentType = ContentType
            request.ContentLength = Postdata.Length

            Dim swRequestWriter As StreamWriter = New StreamWriter(request.GetRequestStream())

            If jsonDataBytes <> "" Then '.Length
                swRequestWriter.Write(jsonDataBytes, 0, (Encoding.UTF8.GetBytes(jsonDataBytes)).Length)
            End If


            If Accept <> "" Then
                request.Accept = Accept
            End If

            If AcceptEncoding <> "" Then
                request.Headers.Add("Accept-Encoding", AcceptEncoding)
            End If
            If AutomaticDecompression = "1" Then
                request.AutomaticDecompression = DecompressionMethods.GZip
            ElseIf AutomaticDecompression = "2" Then
                request.AutomaticDecompression = DecompressionMethods.Deflate
            ElseIf AutomaticDecompression = "0" Then
                request.AutomaticDecompression = DecompressionMethods.None
            End If



            swRequestWriter.Write(Postdata)
            swRequestWriter.Close()
            Try
                response = request.GetResponse()
            Catch ex As WebException
                If documentCode401 Then
                    Dim documentoNaoAutorizado As New HtmlDocument
                    Dim responseStream = ex.Response.GetResponseStream
                    Dim readerStream = New StreamReader(responseStream, Encoding.UTF8)
                    Dim htmlDocumento = HttpUtility.HtmlDecode(readerStream.ReadToEnd().Trim())
                    documentoNaoAutorizado.OptionFixNestedTags = CorrigirHTML
                    documentoNaoAutorizado.LoadHtml(htmlDocumento)
                    Location = ex.Response.ResponseUri.ToString
                    If ex.Response.Headers.GetValues(CookieName) IsNot Nothing Then
                        Cookie = ex.Response.Headers.GetValues(CookieName).First
                    End If

                    Return documentoNaoAutorizado
                End If
            End Try

            Me.response = response

            Dim srResponseReader As StreamReader
            If Encoding Is Nothing Then
                srResponseReader = New StreamReader(response.GetResponseStream(), Encoding.UTF8)
            Else
                srResponseReader = New StreamReader(response.GetResponseStream(), Encoding)
            End If

            Dim html = HttpUtility.HtmlDecode(srResponseReader.ReadToEnd().Trim)

            If Not String.IsNullOrEmpty(response.GetResponseHeader(CookieName)) Then
                Cookie = response.GetResponseHeader(CookieName)
            End If

            Location = response.ResponseUri.ToString
            Me.response = response

            srResponseReader.Close()

            Dim documento = New HtmlDocument()
            documento.OptionFixNestedTags = CorrigirHTML
            documento.LoadHtml(html)

            Return documento
        End Function


        Public Sub AddPostData(ByVal key As String, ByVal val As String, Optional ByVal ultimoChar As Boolean = True)

            If ultimoChar Then
                If Encoding IsNot Nothing Then
                    Postdata &= HttpUtility.UrlEncode(key, Encoding) & "=" & HttpUtility.UrlEncode(val, Encoding) & "&"
                Else
                    Postdata &= HttpUtility.UrlEncode(key, Encoding.UTF8) & "=" & HttpUtility.UrlEncode(val, Encoding.UTF8) & "&"
                End If
            Else
                If Encoding IsNot Nothing Then
                    Postdata &= HttpUtility.UrlEncode(key, Encoding) & "=" & HttpUtility.UrlEncode(val, Encoding)
                Else
                    Postdata &= HttpUtility.UrlEncode(key, Encoding.UTF8) & "=" & HttpUtility.UrlEncode(val, Encoding.UTF8)
                End If
            End If

        End Sub

        Public Function CriarPostJsonRequestJson() As JObject

            Dim request As WebRequest

            request = WebRequest.Create(Url)
            request.ContentLength = jsonDataBytes.Length
            request.ContentType = ContentType
            request.Method = "POST"

            If Not String.IsNullOrEmpty(Cookie) Then
                request.Headers.Add(HttpRequestHeader.Cookie, Cookie)
            End If

            If TimeOut <> 0 Then
                request.Timeout = TimeOut
            End If

            request.Headers.Add(HttpRequestHeader.AcceptLanguage, Lang)

            If UpgradeInsecureRequests <> "" Then
                request.Headers.Add("Upgrade-Insecure-Requests", UpgradeInsecureRequests)
            End If

            If Captcha <> "" Then
                request.Headers.Add("captcha", Captcha)
            End If

            If CacheControl <> "" Then
                request.Headers.Add("Cache-Control", CacheControl)
            End If
            If AcceptTransferEncoding <> "" Then
                request.Headers.Add("Accept-Transfer-Encoding", AcceptTransferEncoding)
            End If



            Using requestStream = request.GetRequestStream
                requestStream.Write(Encoding.UTF8.GetBytes(jsonDataBytes), 0, Encoding.UTF8.GetBytes(jsonDataBytes).Length)
                requestStream.Close()




                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        Return JObject.Parse(reader.ReadToEnd.Trim)
                    End Using
                End Using
            End Using

        End Function

    End Class
End Namespace









