Imports HtmlAgilityPack
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports System.Net
Imports System.IO
Imports OpenQA.Selenium.Support.UI

Public Class Form1
    Public itemToSearch As String = ""
    Dim numberOfPages As Integer = 0
    Public Class ItemResult
        Public Property item_name As String
        Public Property Price As String
        Public Property Location As String
        Public Property URL As String
        Public Property Image As String
    End Class
    Public Class Example
        Public Property item_results As ItemResult()

    End Class

    Public queryTimeout As Integer = 10
    Private Async Function scrollToBottom(driver As ChromeDriver) As Task
        While True
            Dim currentYPosition = driver.ExecuteScript("return window.pageYOffset;")
            'Scroll by a bit.
            driver.ExecuteScript("window.scrollTo(0, window.scrollY + 500)")

            'Wait for stuff to load.
            Await Task.Delay(200)

            'Calculate new scroll height and compare it with last height.
            Dim newHeight = driver.ExecuteScript("return window.pageYOffset;")
            If newHeight = currentYPosition Then 'This means the browser reached the max height.
                Console.WriteLine("Reached bottom of the page.")
                lblStatus.Text = "Data rendered from Shopee."
                Exit While
            End If
        End While
    End Function
    'Sample Image: https://cf.shopee.ph/file/c54850b68337e0cce9b41780867613eb_tn
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()
    End Sub
    Public Sub disableAllControls()
        For Each c As Control In Me.Controls
            c.Enabled = False
        Next
    End Sub
    Public Sub enableAllControls()
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next
    End Sub
    Private Async Function btnSubmit_ClickAsync(sender As Object, e As EventArgs) As Task Handles btnSubmit.Click
        disableAllControls()
        itemToSearch = tbSearch.Text
        Dim resultsDataTable = Await SearchShopee(tbSearch.Text)
        DataGridView1.DataSource = resultsDataTable
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        DataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
        Dim btn As New DataGridViewButtonColumn()
        DataGridView1.Columns.Add(btn)
        btn.HeaderText = "Item Links"
        btn.Text = "View on browser"
        btn.Name = "btnLink"
        btn.UseColumnTextForButtonValue = True

        Dim btnAddtoPR As New DataGridViewButtonColumn()
        DataGridView1.Columns.Add(btnAddtoPR)
        btnAddtoPR.HeaderText = "Add to PR"
        btnAddtoPR.Text = "Add to Purchase Request"
        btnAddtoPR.Name = "btnAddToPR"
        btnAddtoPR.UseColumnTextForButtonValue = True

        'hide product page
        DataGridView1.Columns("Product Page").Visible = False
        DataGridView1.Columns("Image URL").Visible = False

        DataGridView1.AllowUserToAddRows = False
        enableAllControls()
    End Function

    Private Async Function SearchShopee(ByVal itemToSearch As String) As Task(Of DataTable)

        DataGridView1.Columns.Clear()
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        Dim resultsTable As New DataTable
        resultsTable.Columns.Add("Item Name")
        resultsTable.Columns.Add("Image", GetType(Image))
        resultsTable.Columns.Add("Location")
        resultsTable.Columns.Add("Product Page")
        resultsTable.Columns.Add("Prices")
        resultsTable.Columns.Add("Image URL")
        Dim encodedItemString As String = ""
        For Each character In itemToSearch
            If character = " " Then
                encodedItemString += "%20"
            Else
                encodedItemString += character
            End If
        Next
        Dim html As String
        'Use Selenium to browse using chrome.
        'set options to make chrome headless.
        Dim options As New ChromeOptions()
        Dim optionParamaters As String() = {"incognito", "nogpu", "disable-gpu", "no-sandbox", "window-size=1280,1280", "enable-javascript"}
        options.AddArguments(optionParamaters)

        'declare a chromedriver service
        Dim service = ChromeDriverService.CreateDefaultService()
        service.HideCommandPromptWindow = True
        'attach the option to the new driver manager.
        Dim driver = New ChromeDriver(service, options)
        'driver.Manage().Window.Maximize()
        While True
            'Go to the search result url
            Try
                lblStatus.Text = "Waiting for a response from Shopee.."
                driver.Navigate().GoToUrl($"https://shopee.ph/search?keyword={encodedItemString}")
                'set driver to wait for 10 seconds for the page to load.
                Dim wait As WebDriverWait = New WebDriverWait(driver, TimeSpan.FromSeconds(10))
                wait.IgnoreExceptionTypes(GetType(NoSuchElementException))
                'wait.Until(Function(webDriver)
                '               Return webDriver.FindElement(By.ClassName("shopee-search-item-result__items"))
                '           End Function)
                wait.Until(Function(webDriver)
                               Try
                                   driver.FindElement(By.ClassName("shopee-search-item-result__items"))
                               Catch ex As Exception
                                   Dim exType As Type = ex.GetType()
                                   If exType = GetType(NoSuchElementException) Or exType = GetType(InvalidOperationException) Then
                                       Return False
                                       'by returning false, wait will still rerun the function.
                                   Else
                                       Throw
                                   End If
                               End Try
                               Return True
                           End Function)
                lblStatus.Text = "Shopee items found!"
                Exit While
            Catch ex As Exception
                Dim dlgrslt = MessageBox.Show($"Exception occured when loading webpage.{Environment.NewLine}{ex.Message}", "Oops, something went wrong!", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
                If dlgrslt = DialogResult.Retry Then
                    driver.Close()
                Else
                    driver.Close()
                    Application.Restart()
                End If
            End Try
        End While

        'Scroll up to the bottom of the page to load all items.
        lblStatus.Text = "Enabling scripts from shopee.."
        Await scrollToBottom(driver)
        'set the pagesource to the html string variable.
        lblStatus.Text = "Parsing results.."
        html = driver.PageSource
        driver.Close()

        'Load it as an HTML Document.
        Try
            Dim htmlDocument As New HtmlDocument
            htmlDocument.LoadHtml(html)

            'Make a single node out of a body.
            Dim htmlBody = htmlDocument.DocumentNode.SelectSingleNode("//body")

            'find out the number of page in result

            'Select the search result node.
            Dim searchResultsNode As HtmlNode = htmlBody.SelectSingleNode(".//div[(@class='row shopee-search-item-result__items')]")
            'Save the search result in an HTMLNodeCollection.
            ProgressBar1.Maximum = searchResultsNode.ChildNodes.Count
            For Each items As HtmlNode In searchResultsNode.SelectNodes(".//div[(@class='col-xs-2-4 shopee-search-item-result__item')]")
                ProgressBar1.Increment(1)
                lblStatus.Text = $"Parsing results.. ({ProgressBar1.Value}/{ProgressBar1.Maximum})"
                Dim itemName As String = items.SelectSingleNode(".//div[(@class='ie3A+n bM+7UW Cve6sh')]").InnerText
                'Dim price As String = items.SelectSingleNode(".//div[(@class='ZEgDH9')]").InnerText
                Dim location As String = items.SelectSingleNode(".//div[(@class='zGGwiV')]").InnerText
                Dim url As String = items.SelectSingleNode(".//a[(@data-sqe='link')]").Attributes("href").Value
                Dim imageUrl As String = items.SelectSingleNode(".//img[(@class='_7DTxhh vc8g9F')]").Attributes("src").Value
                Dim tClient As New WebClient
                Dim tImage = New Bitmap(Bitmap.FromStream(New MemoryStream(tClient.DownloadData(imageUrl))), 120, 120)
                Dim price As String = ""
                'Price varies, sometimes shopee displays the original price, or sometimes it displays a price range
                'And sometimes it also displays the old price, and the new price beside it. So we'll capture the NODE
                'of the whole price section instead and count the number of children inside it.
                Dim priceNode As HtmlNode = items.SelectSingleNode(".//div[(@class='hpDKMN')]")
                'this is either a Price Range or a Single Price
                Dim currentPriceClass As HtmlNode = priceNode.SelectSingleNode(".//div[(@class='vioxXd rVLWG6')]")
                For Each span As HtmlNode In currentPriceClass.ChildNodes
                    If span.Name = "span" Then
                        price += span.InnerText
                    Else
                        price += " - "
                    End If
                Next
                Dim newRow = resultsTable.NewRow()
                newRow.Item("Item Name") = itemName
                newRow.Item("Image") = tImage
                newRow.Item("Location") = location
                newRow.Item("Product Page") = "https://www.shopee.ph" + url
                newRow.Item("Prices") = price
                'newRow.Item("Image URL") = imageUrl

                resultsTable.Rows.Add(newRow)
                'resultsTable.Columns.Add("Item Name")
                'resultsTable.Columns.Add("Location")
                'resultsTable.Columns.Add("Product Page")
                'resultsTable.Columns.Add("Prices")
                'resultsTable.Columns.Add("Image URL")
            Next
            ProgressBar1.Value = ProgressBar1.Maximum
            ProgressBar1.Visible = False
            lblStatus.Text = "Fetch complete!"
        Catch ex As Exception
            Dim result = MessageBox.Show($"Error occured when parsing results.{Environment.NewLine}{ex.Message}", "Oops, something went wrong!", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
            If result = DialogResult.Retry Then
                Application.Restart()
            Else
                Me.Close()
            End If
        End Try
        Return resultsTable
    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim senderGrid = DirectCast(sender, DataGridView)
        If TypeOf senderGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView1.Columns("btnLink").Index Then
                Console.WriteLine("Link button was pressed.")
                Process.Start(DataGridView1.Rows(e.RowIndex).Cells("Product Page").Value.ToString)
            Else
                Console.WriteLine("add to pr button was pressed.")
            End If
        End If
    End Sub

    Private Async Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

    End Sub

    'results': '//div[contains(@class, 'shopee-search-item-result__items')]'
    ','body': '//div[contains(@class, \'article\')]'

End Class
