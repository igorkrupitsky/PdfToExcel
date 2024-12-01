Imports System.Net.Http
Imports System.Net.Http.Headers

Public Class AICaller

    Public sOpenAiKey As String = ""
    Public sAnthropicApiKey As String = ""
    Public sModel As String = ""

    Public Function SendImg(sImagePath As String, sPrompt As String, sDetail As String) As String
        If L(sModel, 4) = "gpt-" Then
            Return SendOpenAiImg(sImagePath, sPrompt, sDetail)
        Else
            Return SendAnthropicImg(sImagePath, sPrompt)
        End If
    End Function

    Public Function SendAnthropicImg(sImagePath As String, sPrompt As String) As String
        'https://docs.anthropic.com/en/docs/build-with-claude/vision#how-to-use-vision
        Dim sUrl As String = "https://api.anthropic.com/v1/messages"
        Dim iMaxTokens As Integer = 4096
        Dim dTemperature As Double = 0
        Dim data As String = ""

        If IO.File.Exists(sImagePath) Then
            Dim image_data As String = GetFile64(sImagePath)
            data = "{" &
                    """model"": """ & sModel & """," &
                    """max_tokens"": " & iMaxTokens & "," &
                    """messages"": [{" &
                    """role"":""user""," &
                    """content"": [" &
                    "{""type"": ""image"",""source"": {""type"": ""base64"",""media_type"": ""image/jpeg"",""data"": """ & image_data & """}}," &
                    "{""type"": ""text"",""text"": """ & PadQuotes(sPrompt) & """}" &
                    "]}]}"
        Else
            data = "{"
            data += """model"": """ & sModel & ""","
            data += """messages"": [{""role"":""user"", ""content"": """ & PadQuotes(sPrompt) & """}],"
            data += """system"": ""You are Claude, an AI assistant created by Anthropic to be helpful, harmless, and honest."","
            data += """max_tokens"": " & iMaxTokens & ","
            data += """temperature"": " & dTemperature
            data += "}"
        End If

        Dim oHeaders As New Hashtable
        oHeaders("x-api-key") = sAnthropicApiKey
        oHeaders("anthropic-version") = "2023-06-01"
        Dim sJson As String = SendHttpRequest(sUrl, data, oHeaders)

        Dim oJavaScriptSerializer As New System.Web.Script.Serialization.JavaScriptSerializer
        Dim oJson As Hashtable = oJavaScriptSerializer.Deserialize(Of Hashtable)(sJson)
        Return oJson("content")(0)("text")
    End Function

    Public Function SendOpenAiImg(sImagePath As String, sPrompt As String, sDetail As String) As String
        'short side of the image should be less than 768px and the long side should be less than 2,000px.
        'https://platform.openai.com/docs/guides/vision
        Const sUrl As String = "https://api.openai.com/v1/chat/completions"
        Dim data As String = ""

        If sImagePath <> "" AndAlso IO.File.Exists(sImagePath) Then
            Dim image_data As String = GetFile64(sImagePath)
            data = "{" &
                """model"": """ & sModel & """," &
                """messages"": [{" &
                """role"":""user""," &
                """content"": [" &
                "{""type"": ""image_url"",""image_url"": {""url"": ""data:image/jpeg;base64," & image_data & """,""detail"": """ & sDetail & """}}," &
                "{""type"": ""text"",""text"": """ & PadQuotes(sPrompt) & """}" &
                "]}]}"
        Else
            data = "{" &
            " ""model"":""" & sModel & """," &
            " ""messages"": [{""role"":""user"", ""content"": """ & PadQuotes(sPrompt) & """}]" &
            "}"
        End If

        Dim oHeaders As New Hashtable
        oHeaders("Authorization") = sOpenAiKey
        Dim sJson As String = SendHttpRequest(sUrl, data, oHeaders)

        Dim oJavaScriptSerializer As New System.Web.Script.Serialization.JavaScriptSerializer
        Dim oJson As Hashtable = oJavaScriptSerializer.Deserialize(Of Hashtable)(sJson)
        Return oJson("choices")(0)("message")("content")
    End Function

    Private Function GetFile64(imagePath As String) As String
        Dim fileBytes() As Byte = IO.File.ReadAllBytes(imagePath)
        Return Convert.ToBase64String(fileBytes).Replace(vbCrLf, "").Replace(vbCr, "").Replace(vbLf, "")
    End Function

    Private Function SendHttpRequest(url As String, data As String, oHeaders As Hashtable) As String
        Using httpClient As New HttpClient()
            ' Create the request
            Dim request As New HttpRequestMessage(HttpMethod.Post, url)

            ' Set the content of the request
            request.Content = New StringContent(data, Text.Encoding.UTF8, "application/json")

            ' Add headers to the request
            For Each oKey As DictionaryEntry In oHeaders
                If oKey.Key.ToString() = "Authorization" Then
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", oKey.Value.ToString())
                Else
                    request.Headers.Add(oKey.Key.ToString(), oKey.Value.ToString())
                End If
            Next

            httpClient.Timeout = TimeSpan.FromSeconds(60 * 30)

            ' Send the request
            Dim response As HttpResponseMessage = httpClient.SendAsync(request).GetAwaiter().GetResult()

            ' Ensure the request was successful
            response.EnsureSuccessStatusCode()

            ' Read the response content
            Dim responseBody As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            Return responseBody
        End Using
    End Function

    Private Function PadQuotes(s As String) As String
        s = s.Replace("\", "\\")
        s = s.Replace(vbCrLf, "\n")
        s = s.Replace(vbCr, "\r")
        s = s.Replace(vbLf, "\f")
        s = s.Replace(vbTab, "\t")
        Return s.Replace("""", "\""")
    End Function

    Private Function L(s As String, i As Integer) As String
        Return Microsoft.VisualBasic.Left(s, i)
    End Function

End Class
