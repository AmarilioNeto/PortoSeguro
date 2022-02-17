Imports System.Net
Imports System.Net.Security
Imports System.Net.HttpWebRequest
Imports System.Text.RegularExpressions
Imports System.Text
Imports Port_Seguro_API.Acesso
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Net.WebRequest



Module AcessoApi


    Public Sub main()
        Dim test As New API
        test.AcessoPrincipal()
    End Sub

    Public Class API
        Private host As String = "wws.averbeporto.com.br"
        Public Function AcessoPrincipal()
            Dim request = New PortoSeguro With {
            .Host = host,
            .Url = "https://wws.averbeporto.com.br/websys/php/conn.php",
            .Accept = "*/*",
            .AcceptEncoding = "gzip, deflate, br",
            .Lang = "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            .userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0",
            .XMLHttpRequest = "XMLHttpRequest",
            .Cookie = "cadba64164ee79cf1aaa70361d7abedf",
            .ContentType = "application/x-www-form-urlencoded; charset=UTF-8",
            .ContentLength = "19",
            .Referer = "https://wws.averbeporto.com.br/websys/?comp=5",
            .Origin = "https://wws.averbeporto.com.br",
            .AutomaticDecompression = "1"
            }

            request.AddPostData("mod", "login")
            request.AddPostData("comp", "5")


            Dim documento = request.CriarPostHttpRequestAsDocument()
            Dim cookie = request.Cookie


            request = New PortoSeguro With {
     .Host = host,
    .Url = "https://wws.averbeporto.com.br/websys/php/conn.php",
    .userAgent = “Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0”,
    .Accept = "*/*",
    .AcceptEncoding = "gzip, deflate, br",
    .Lang = "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    .ContentLength = "52",
    .ContentType = "application/x-www-form-urlencoded; charset=UTF-8",
    .Referer = "https://wws.averbeporto.com.br/index.html?comp=5",
    .Cookie = cookie,
    .XMLHttpRequest = "XMLHttpRequest",
    .Origin = "https://wws.averbeporto.com.br",
    .dump = "2"
    }
            request.AddPostData("mod", "login")
            request.AddPostData("comp", "5")
            request.AddPostData("user", "57460644000180")
            request.AddPostData("pass", "Beto@134")


            documento = request.CriarPostHttpRequestAsDocument()
            cookie = request.Cookie

            Return documento


        End Function
    End Class
End Module



