Imports RestSharp
Imports Newtonsoft.Json
Imports Squirrel
Imports System.Threading.Tasks
Public Class login
    Public prVersion As String = ""
    'START of Variable Declaration
    Dim resultsList As List(Of Object)
    Public watch As Stopwatch
    Public maxErrorCount As Integer = 30

    Public queryTimeOut As Integer = 15000
    Public titaVersion As String = LasermetPR.My.Application.Info.Version.ToString
    Public allTasks As Root
    Public elapsedTime As Integer
    Public loadDelay As Integer
    Public namesList As New List(Of Object)

    'Stores no. of minutes since last update sent to Monday.com
    Public howLong

    'Task categories
    Public taskCategories() As String = {"Show All", "R&D", "Jobs", "Admin", "Electronics R&D", "Mechanical R&D", "Enclosure", "Systems Designs", "Small Batch Manufacturing"}

    'Variable for the ID of the current log
    Public currentID As String
    Public fSurname As String
    Public fFirstName As String
    Public mondayID As String
    Public department As String
    Public manualLogInID As String
    Public accountItemID As String

    'variables for Switch forms
    Public currentTask As String
    Public currentSubTask As String
    Public currentTimeIn As String
    Public currentProjectNumber As String
    Public accounts As New Root
    'END of Variable Declaration

    'START of Class Declaration for Serialization (Changing ColumnValues for Previous Log)
    Public Class ColumnValuesToChange
        Public Property text_1 As String ' START_Surname
        Public Property dup__of_time_in As String 'Timeout
        Public Property text64 As String 'TiTA Version
    End Class
    'End of Class Declaration for Serialization (Changing ColumnValues for Previous Log)

    'Start Root Classes Declaration
    Public Class Group
        Public Property id As String
        Public Property items As Item()
    End Class
    Public Class Subitem
        Public Property name As String
    End Class
    Public Class ColumnValue
        Public Property title As String
        Public Property text As String
    End Class
    Public Class Item
        Public Property name As String
        Public Property subitems As Subitem()
        Public Property column_values As ColumnValue()
        Public Property id As String
        Public Property group As Group
    End Class
    Public Class Board
        Public Property items As Item()
        Public Property groups As Group()
    End Class
    Public Class Data
        Public Property boards As Board()
    End Class
    Public Class Root
        Public Property data As Data
        Public Property account_id As Integer
    End Class
    'END Root Classes Declaration
    Private Async Sub login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        disableControls()
        Me.TopMost = True
        Me.CenterToParent()
        Me.Text = $"Lasermet PR v{prVersion}"
        cbUsername.AutoCompleteSource = AutoCompleteSource.ListItems
        Dim fetchAccountQuery As String = "query{
                boards(ids:3428362986){
                    items{
                        id    
                        name
                        column_values{
                            title
                            text
                        }
                    }
                }
            }"
        Dim maxRetries = 5
        Dim retries = 0
        Dim myJsonResponse As Root
        'SEND request to monday.com
        While True
            Try
                While True
                    Dim result As Object = Await SendMondayRequestVersion2(fetchAccountQuery)
                    If retries <> maxRetries And retries < maxRetries Then
                        If result(0) = "error" Then
                            'something went wrong. loop wil retry.
                            retries += 1
                        ElseIf result(0) = "success" Then
                            myJsonResponse = JsonConvert.DeserializeObject(Of Root)(result(1))
                            For Each users In myJsonResponse.data.boards(0).items

                                namesList.Add({users.name, users.column_values(0).text, users.column_values(1).text, users.column_values(2).text, users.column_values(3).text})
                                accounts = myJsonResponse
                            Next
                            populateUsersComboBox()
                            Exit While
                        Else
                            retries += 1
                        End If
                    Else
                        Throw New Exception("Program has reached the max number of retries.")
                    End If
                End While
                Exit While
            Catch ex As Exception
                Dim dlgResult = MessageBox.Show(ex.Message, "Oops, something went wrong!", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
                If dlgResult = DialogResult.Retry Then
                    'retry.
                Else
                    Application.Restart()
                End If
            End Try
        End While
        enableControls()
    End Sub
    Public Sub populateUsersComboBox()
        For Each users In namesList
            cbUsername.Items.Add(users(0))
        Next

        If My.Settings.recentUser = "" Then
            cbUsername.Text = "Select User"
        Else
            cbUsername.SelectedItem = My.Settings.recentUser
        End If
    End Sub
    Public Function checkAccountDetails(ByVal surname As String, ByVal password As String, ByVal accounts As Root) As Boolean
        '0 - First Name
        '1 - Monday ID
        '2 - Password
        '3 - Department

        For Each x In accounts.data.boards(0).items
            If x.name = surname Then
                'account found.
                If x.column_values(2).text = password Then
                    'account verified
                    'save all account details in a global variable.
                    fSurname = x.name
                    fFirstName = x.column_values(0).text
                    mondayID = x.column_values(1).text
                    department = x.column_values(3).text
                    accountItemID = x.id
                    Return True
                Else
                    Return False
                End If
            End If
        Next
    End Function
    Public Sub enableControls()
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next
    End Sub
    Public Sub disableControls()
        For Each c As Control In Me.Controls
            c.Enabled = False
        Next
    End Sub
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If checkAccountDetails(cbUsername.SelectedItem, tbPassword.Text, accounts) Then
            'success
            MessageBox.Show($"Welcome, {fFirstName} {fSurname}!", "Login Success!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            My.Settings.recentUser = cbUsername.SelectedItem
            My.Settings.Save()
            Me.Hide()
            OctoPart_API.Show()
        Else
            MessageBox.Show("Username and password does not match. Please try again!", "Oops, something went wrong!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            tbPassword.Clear()
            tbPassword.Select()
        End If
    End Sub

    Public Async Function SendMondayRequestVersion2(ByVal myQuery As String) As Task(Of Object)
        Dim options = New RestClientOptions("https://api.monday.com/v2")
        options.ThrowOnAnyError = True
        options.MaxTimeout = queryTimeOut
        Dim client = New RestClient(options)
        Dim request = New RestRequest()
        request.Timeout = queryTimeOut
        request.Method = Method.Post
        request.AddHeader("Authorization", "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjE1MjU2NzQ3OCwidWlkIjoxNTA5MzQwNywiaWFkIjoiMjAyMi0wMy0yNVQwMTo0Njo1My4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjYxMjMxMCwicmduIjoidXNlMSJ9.F1TqwLS-QsuM8Ss3UcgskbNFUIer1dfwfoLyPMq1pbc")
        request.AddQueryParameter("query", myQuery)
        Dim response = New RestResponse
        response = Await client.PostAsync(request)
        'If response.IsSuccessStatusCode = True Then
        '    Return response.Content
        'Else
        '    Return False
        'End If
        If response.IsSuccessStatusCode Then
            'response has a statuscode of 200
            'but it might have a parse error, which still is status 200.
            If response.Content.Contains("error") Or response.Content.Contains("error_message") Or response.Content.Contains("errors") Then
                'response has a status code 200, but has a monday.com error.
                Return {"error", response.Content}
            Else
                'response has a status code 200, with readable results.
                Return {"success", response.Content}
            End If
        Else
            Throw New System.Exception("An error has occured at function: SendMondayRequestVersion2")
        End If
    End Function

End Class