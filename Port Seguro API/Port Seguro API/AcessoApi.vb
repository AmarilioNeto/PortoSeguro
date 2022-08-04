Imports System.Net
Imports System.Net.Security
Imports System.Net.HttpWebRequest
Imports System.Text.RegularExpressions
Imports System.Text
Imports Port_Seguro_API.Acesso
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Net.WebRequest
Imports System.Net.Http
Imports System.Web
Imports HtmlAgilityPack
Imports System.Collections.Specialized

Module AcessoApi


    Public Sub main()
        Dim test As New API
        test.AcessoPrincipal()
    End Sub

    Public Class API
        Private host As String = "wws.averbeporto.com.br"
        Public Function AcessoPrincipal()

            'primeiro acesso Login
            Dim request = New PortoSeguro With {
            .Host = host,
            .Url = "https://apis.averbeporto.com.br/php/conn.php",
            .userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0",
            .ContentType = "application/x-www-form-urlencoded"
            }
            request.AddPostData("mod", "*****")
            request.AddPostData("comp", "***")
            request.AddPostData("user", "*****")
            request.AddPostData("pass", "*****")

            Dim documento = request.CriarPostHttpRequestAsDocument()
            Dim cookie = request.Cookie

            'segundo acesso requsisçao e envio de arquivo
            Dim postData = New NameValueCollection
            postData.Add("comp", "5")
            postData.Add("mod", "Upload")
            postData.Add("path", "eguarda/php/")
            postData.Add("recipient", "")

            documento = request.CriarPostHttpRequestAsDocument()
            cookie = request.Cookie



            Return documento


        End Function

    End Class
End Module



