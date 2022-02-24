Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Http


Public Class UploadController
    Inherits ApiController

    Public Async Function PostFormData() As Task(Of HttpResponseMessage)

        Dim root As String = HttpContext.Current.Server.MapPath("~/App_Data")
        Dim provider = New MultipartFormDataStreamProvider(root)

        Try
            Await Request.Content.ReadAsMultipartAsync(provider)

            For Each file As MultipartFileData In provider.FileData
                Trace.WriteLine(file.Headers.ContentDisposition.FileName)
                Trace.WriteLine("Server file path: " & file.LocalFileName)
            Next

            Return Request.CreateResponse(HttpStatusCode.OK)
        Catch e As System.Exception
            Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e)
        End Try
    End Function
End Class
