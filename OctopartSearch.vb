Imports Newtonsoft.Json
Imports System.Text
Imports Nexar.Client
Imports Nexar.Client.Token
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks

Module OctopartSearch

    Private NexarToken As String
    Private tokenLife As TimeSpan = TimeSpan.FromDays(1)
    Private tokenExpiry As DateTime

    Public responseObj
    Dim mpn As String = "5EFM1S"
    Dim partKeyword As String = "5000PF"
    Dim httpClient As HttpClient
    Dim queryTimeout = 15000

    Public Class Request

        Public Property query As String
        Public Property variables As Dictionary(Of String, String)

    End Class

    'Classes for Nexar Part Search Results

    Public Class BestImage
        Public Property url As String
    End Class

    Public Class Child
        Public Property id As String
        Public Property name As String
        Public Property path As String
        Public Property numParts As Integer
        Public Property parentId As String
    End Class

    Public Class Category
        Public Property name As String
        Public Property id As String
    End Class

    Public Class Manufacturer
        Public Property name As String
        Public Property id As String
    End Class

    Public Class MedianPrice1000
        Public Property convertedPrice As Double
        Public Property convertedCurrency As String
        Public Property quantity As Integer
    End Class

    Public Class Company
        Public Property name As String
        Public Property homepageUrl As String
    End Class

    Public Class Price
        Public Property convertedPrice As Double
        Public Property convertedCurrency As String
        Public Property quantity As Integer
    End Class

    Public Class Offer
        Public Property clickUrl As String
        Public Property inventoryLevel As Integer
        Public Property moq As Integer?
        Public Property prices As Price()
        Public Property packaging As String
    End Class

    Public Class Seller
        Public Property company As Company
        Public Property isAuthorized As Boolean
        Public Property offers As Offer()
    End Class

    Public Class Part
        Public Property mpn As String
        Public Property bestImage As BestImage
        Public Property shortDescription As String
        Public Property manufacturer As Manufacturer
        Public Property medianPrice1000 As MedianPrice1000
        Public Property category As Category
        Public Property sellers As Seller()
    End Class

    Public Class Result
        Public Property part As Part
    End Class

    Public Class SupSearch
        Public Property hits As Integer
        Public Property results As Result()
    End Class

    Public Class SupCategory
        Public Property id As String
        Public Property name As String
        Public Property path As String
        Public Property numParts As Integer
        Public Property parentId As String
        Public Property children As Child()
    End Class

    Public Class Data
        Public Property supSearch As SupSearch
        Public Property supCategories As SupCategory()
    End Class

    Public Class Extensions
        Public Property requestId As String
    End Class

    Public Class NexarAPIResponse
        Public Property data As Data
        Public Property extensions As Extensions
    End Class

    'Classes for error handling

    Public Class Remote
        Public Property message As String
    End Class

    Public Class _Extensions
        Public Property remote As Remote
        Public Property schemaName As String
    End Class

    Public Class _Error
        Public Property message As String
        Public Property extensions As _Extensions
    End Class

    Public Class ErrorMsg
        Public Property errors As _Error()
        Public Property extensions As Extensions
    End Class

    Private Async Function updateToken() As Task(Of Object)

        Dim clientID As String = "11b22d9d-3b47-4ebe-87a8-d59ff12a678c"
        Dim clientSecret As String = "fe05c906-9d0e-4f22-aedc-aabfdf5d05dc"
        httpClient = New HttpClient()

        If String.IsNullOrEmpty(NexarToken) Or DateTime.UtcNow() >= tokenExpiry Then
            tokenExpiry = DateTime.Now() + tokenLife
            Dim authclient = New HttpClient()
            NexarToken = Await authclient.GetNexarTokenAsync(clientID, clientSecret)
        End If


        httpClient.BaseAddress = New Uri("https://api.nexar.com/graphql")

        httpClient.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Bearer", NexarToken)

    End Function


    Public Async Function sendAPIRequest(ByVal query As String) As Task(Of Object)


        Await updateToken()
        Dim request = New Request

        request.query = query

        Dim requestString As String = JsonConvert.SerializeObject(request)
        Dim httpResponse As HttpResponseMessage = Await httpClient.PostAsync(httpClient.BaseAddress, New StringContent(requestString, Encoding.UTF8, "application/json"))
        Dim responseString As String = Await httpResponse.Content.ReadAsStringAsync()

        Return responseString

    End Function


End Module
