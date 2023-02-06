Imports HtmlAgilityPack
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports System.Net
Imports System.IO
Imports OpenQA.Selenium.Support.UI

Public Class Form1
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
            driver.ExecuteScript("window.scrollTo(0, window.scrollY + 150);")

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
        'Dim tClient As WebClient = New WebClient
        'Dim tImage = New Bitmap(Bitmap.FromStream(New MemoryStream(tClient.DownloadData("https://cf.shopee.ph/file/c54850b68337e0cce9b41780867613eb_tn"))), 120, 120)

        ''This is for test
        'Dim datatable As New DataTable
        'datatable.Columns.Add("Product ID")
        'datatable.Columns.Add("Image", GetType(Image))
        'datatable.Columns.Add("Short Description")
        'datatable.Columns.Add("Location")
        'datatable.Columns.Add("Price")


        ''populate fields
        'Dim datarow = datatable.NewRow()
        'datarow.Item("Product ID") = "Hello World"
        'datarow.Item("Short Description") = "This is a sample short description. Lores ipsum lores ipsum lores ipsum generator words goes here."
        'datarow.Item("Location") = "Cebu City"
        'datarow.Item("Image") = tImage
        'datarow.Item("Price") = "P100"

        'datatable.Rows.Add(datarow)

        'Dim datarow1 = datatable.NewRow()
        'datarow1.Item("Product ID") = "Hello World123"
        'datarow1.Item("Short Description") = "This i12312312s a sample short description. Lores ipsum lores ipsum lores ipsum generator words goes here."
        'datarow1.Item("Location") = "Cebu 123123City"
        'datarow1.Item("Price") = "P101230"
        'datarow1.Item("Image") = tImage

        'datatable.Rows.Add(datarow1)

        ''databind table to datagridview
        'DataGridView1.DataSource = datatable
        'DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        'DataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True

        ''add buttons
        'Dim btn As New DataGridViewButtonColumn()
        'DataGridView1.Columns.Add(btn)
        'btn.HeaderText = "Item Links"
        'btn.Text = "View on browser"
        'btn.Name = "btnLink"
        'btn.UseColumnTextForButtonValue = True
    End Sub

    Private Async Function btnSubmit_ClickAsync(sender As Object, e As EventArgs) As Task Handles btnSubmit.Click
        Dim resultsDataTable = Await searchShopee(tbSearch.Text)
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




        'add an image column
    End Function

    Private Async Function searchShopee(ByVal itemToSearch As String) As Task(Of DataTable)
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
        lblStatus.Text = "Rendering data from shopee.."
        'Use Selenium to browse using chrome.
        Dim driver = New ChromeDriver
        While True
            driver.Manage().Window.Maximize()
            'Go to the search result url
            driver.Navigate().GoToUrl($"https://shopee.ph/search?keyword={encodedItemString}")

            Try
                'set driver to wait for 10 seconds for the page to load.
                Dim wait As WebDriverWait = New WebDriverWait(driver, TimeSpan.FromSeconds(10))
                wait.IgnoreExceptionTypes(GetType(NoSuchElementException))
                wait.Until(Function(webDriver)
                               Return webDriver.FindElement(By.ClassName("shopee-search-item-result__items"))
                           End Function)
                Exit While
            Catch ex As Exception
                Dim dlgrslt = MessageBox.Show($"Exception occured when loading webpage.{Environment.NewLine}{ex.Message}", "Oops, something went wrong!", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
                If dlgrslt = DialogResult.Retry Then
                Else
                    Application.Restart()
                End If
            End Try
        End While

        'Scroll up to the bottom of the page to load all items.
        Await scrollToBottom(driver)
        'set the pagesource to the html string variable.
        lblStatus.Text = "Parsing results.."
        html = driver.PageSource
        driver.Close()

        'Load it as an HTML Document.
        Dim htmlDocument As New HtmlDocument
        htmlDocument.LoadHtml(html)

        'Make a single node out of a body.
        Dim htmlBody = htmlDocument.DocumentNode.SelectSingleNode("//body")

        'Select the search result node.
        Dim searchResultsNode As HtmlNode = htmlBody.SelectSingleNode(".//div[(@class='row shopee-search-item-result__items')]")

        'Save the search result in an HTMLNodeCollection.
        For Each items As HtmlNode In searchResultsNode.SelectNodes(".//div[(@class='col-xs-2-4 shopee-search-item-result__item')]")
            Dim itemName As String = items.SelectSingleNode(".//div[(@class='ie3A+n bM+7UW Cve6sh')]").InnerText
            'Dim price As String = items.SelectSingleNode(".//div[(@class='ZEgDH9')]").InnerText
            Dim location As String = items.SelectSingleNode(".//div[(@class='zGGwiV')]").InnerText
            Dim url As String = items.SelectSingleNode(".//a[(@data-sqe='link')]").Attributes("href").Value
            Dim imageUrl As String = items.SelectSingleNode(".//img[(@class='_7DTxhh vc8g9F')]").Attributes("src").Value
            Dim tClient As WebClient = New WebClient
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
            newRow.Item("Product Page") = url
            newRow.Item("Prices") = price
            'newRow.Item("Image URL") = imageUrl

            resultsTable.Rows.Add(newRow)
            'resultsTable.Columns.Add("Item Name")
            'resultsTable.Columns.Add("Location")
            'resultsTable.Columns.Add("Product Page")
            'resultsTable.Columns.Add("Prices")
            'resultsTable.Columns.Add("Image URL")
        Next
        Return resultsTable
    End Function

    'results': '//div[contains(@class, 'shopee-search-item-result__items')]'
    ','body': '//div[contains(@class, \'article\')]'

End Class
