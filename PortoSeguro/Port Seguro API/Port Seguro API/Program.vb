
Imports System.IO
Imports System.Net
Imports System.Xml
Imports System.Text



Module Program
    Sub Main()
        Dim test As New Porto

        test.Capturar()


    End Sub
    Public Class Porto

        Public Function Capturar() As String

            Dim enderecoWebService = " http://api.averbeporto.com.br/php/conn.php (Plain text) - HTTP/2"


            'Requisição para acesso ao webservice de integração
            Dim request As HttpWebRequest = Nothing
            Dim response As HttpWebResponse = Nothing

            request = HttpWebRequest.Create(enderecoWebService)
            request.Method = "POST"
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0"





            Dim rd = New StreamReader(response.GetResponseStream(), Encoding.UTF8)
            Dim documentoString = rd.ReadToEnd.Trim
            ' End If

            If documentoString.StartsWith("<?xml") Or documentoString.StartsWith("<html") Then

                Dim sb As New StringBuilder
                Dim settings As XmlWriterSettings = New XmlWriterSettings()
                settings.Encoding = Encoding.Unicode
                settings.Indent = True

                Using reader As XmlReader = XmlReader.Create(New StringReader(documentoString))


                End Using
            Else

                Throw New Exception(documentoString)

            End If
            Return documentoString
        End Function
    End Class
End Module
