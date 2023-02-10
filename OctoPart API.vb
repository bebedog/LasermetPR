Imports RestSharp
Imports Newtonsoft.Json
Imports System.Net
Imports System.Linq
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop
Imports System.IO
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support.UI
Imports HtmlAgilityPack

Public Class OctoPart_API

    Public categoriesObj
    Public prTable As New DataTable
    Dim searchObj
    Dim csvSavepath As String
    Dim selectedSource As String

    Private Async Function SearchShopee(ByVal itemToSearch As String, ByVal datagridviewName As DataGridView, ByVal progressBar As ToolStripProgressBar, ByVal lblstatus As ToolStripStatusLabel) As Task(Of DataTable)

        datagridviewName.Columns.Clear()
        progressBar.Value = 0
        progressBar.Visible = True
        progressBar.Alignment = ToolStripItemAlignment.Right
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
        Dim optionParamaters As String() = {"incognito", "nogpu", "disable-gpu", "no-sandbox", "window-size=1280,1280", "enable-javascript", "headless"}
        options.AddArguments(optionParamaters)

        'declare a chromedriver service
        Dim service = ChromeDriverService.CreateDefaultService()
        service.HideCommandPromptWindow = True
        'attach the option to the new driver manager.
        Dim driver As ChromeDriver = New ChromeDriver(service, options)
        'driver.Manage().Window.Maximize()
        While True
            'Go to the search result url
            Try
                lblstatus.Text = "Waiting for a response from Shopee.."
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
                lblstatus.Text = "Shopee items found!"
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
        lblstatus.Text = "Enabling scripts from shopee.."
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
                lblstatus.Text = "Data rendered from Shopee."
                Exit While
            End If
        End While
        'set the pagesource to the html string variable.
        lblstatus.Text = "Parsing results.."
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
            progressBar.Maximum = searchResultsNode.ChildNodes.Count
            For Each items As HtmlNode In searchResultsNode.SelectNodes(".//div[(@class='col-xs-2-4 shopee-search-item-result__item')]")
                progressBar.Increment(1)
                lblstatus.Text = $"Parsing results.. ({progressBar.Value}/{progressBar.Maximum})"
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
            progressBar.Value = progressBar.Maximum
            progressBar.Visible = False
            lblstatus.Text = "Fetch complete!"
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

    Private Sub disableControls()
        For Each c As Control In Me.Controls
            If c.Name <> cbSources.Name Then
                If c.Name <> labelSource.Name Then
                    c.Enabled = False
                End If
            End If

        Next
    End Sub

    Private Sub enableControls()
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next
    End Sub

    Private Sub OctoPart_API_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Console.WriteLine(Application.StartupPath)
        cbSources.Items.Add("Shopee")
        cbSources.Items.Add("Octopart")
        groupFilters.Visible = False
        groupFilters.Enabled = False
        btnSearch.Enabled = False
        btnExportPR.Enabled = False
        statusLabel.Text = "Please select a source"
        Me.KeyPreview = True
    End Sub

    Private Async Sub populateCategories()

        Dim fetchCategories As String = "query Categories {
                  supCategories {
                    id
                    name
                    path
                    numParts
                    parentId
                    children {
                      id
                      name
                      path
                      numParts
                      parentId
                    }
                  }}"

        Dim categoriesResultString As String = Await sendAPIRequest(fetchCategories)

        If categoriesResultString.Contains("errors") Then
            Dim errorObj = JsonConvert.DeserializeObject(Of ErrorMsg)(categoriesResultString)
            Dim errorMsg As String = errorObj.errors(0).message

            MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If DialogResult.OK Then
                Application.Exit()
                Exit Sub
            End If

        Else

            categoriesObj = JsonConvert.DeserializeObject(Of NexarAPIResponse)(categoriesResultString)

        End If


        Dim categoriesList As New List(Of String)
        categoriesList.Add("All Categories")

        For Each ctg In categoriesObj.data.supCategories
            categoriesList.Add(ctg.name.ToString)
        Next

        tbKeyword.Enabled = True
        cbCategories.Items.AddRange(categoriesList.ToArray)
        cbCategories.SelectedIndex = 0
        groupFilters.Enabled = True
        btnSearch.Enabled = True

        StatusStrip1.Enabled = True
        statusLabel.Text = "You can start searching by Category or a specific keyword ノ( ゜-゜ノ)"

    End Sub

    Private Async Function searchOcto(ByVal keyword As String, ByVal category As String, ByVal subcategory As String) As Task

        disableControls()
        Dim filters As String

        If category = "All Categories" Then
            filters = ""
        ElseIf subcategory = "N/A" Then
            For Each ctg In categoriesObj.data.supCategories
                If ctg.name = category Then
                    filters = "category_id:[""" + ctg.id + """]"
                End If
            Next
        ElseIf subcategory = "All Subcategories" Then
            Dim filterList As New List(Of String)
            For Each ctg In categoriesObj.data.supCategories
                If ctg.name = category Then
                    For Each s In ctg.children
                        filterList.Add(s.id)
                    Next
                End If
            Next
            filters = "category_id:[""" + String.Join(""", """, filterList.ToArray) + """]"

        Else
            For Each ctg In categoriesObj.data.supCategories
                If ctg.name = category Then
                    For Each subc In ctg.children
                        If subc.name = subcategory Then
                            filters = "category_id:[""" + subc.id + """]"
                        End If
                    Next

                End If
            Next
        End If

        'Remove limit in query prior to actual app release
        Dim searchQuery As String = "query PartSearch {
                              supSearch(
                                q: """ + keyword + """
                                filters: {company_id:[""2401"", ""459"", ""3261"",""2628""]," + filters + "}
                                inStockOnly: true
                                country:""PH""
                                currency: ""PHP""
                                limit: 1
                              ) {
                                hits
                                results {
                                  part {
                                    mpn
                                    bestImage{
                                      url
                                    }
                                    shortDescription
                                    manufacturer {
                                      name
                                      id
                                    }
                                    category{
                                      name
                                      id
                                    }
                                    medianPrice1000{
                                      convertedPrice
                                      convertedCurrency
                                      quantity
                                    }
                                    sellers(authorizedOnly: true) {
                                      company {
                                        name
                                        homepageUrl
                                      }
                                      isAuthorized
                                      offers {
                                        clickUrl
                                        inventoryLevel
                                        moq
                                        prices{
                                          convertedPrice
                                          convertedCurrency
                                          quantity
                                        }
                                        packaging
                                      }
                                    }
                                  }
                                }
                              }
                            }"

        Dim searchResultString As String = Await sendAPIRequest(searchQuery)

        If searchResultString.Contains("errors") Then
            Dim errorObj = JsonConvert.DeserializeObject(Of ErrorMsg)(searchResultString)
            Dim errorMsg As String = errorObj.errors(0).message
            MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If DialogResult.OK Then
                Application.Exit()
                Exit Function
            End If
        Else
            searchObj = JsonConvert.DeserializeObject(Of NexarAPIResponse)(searchResultString)
            If searchObj.data.supSearch.hits = 0 Then
                MessageBox.Show($"No part matches found for {tbKeyword.Text}")
                If DialogResult.OK Then
                    enableControls()
                    Exit Function
                End If
            End If
        End If

        resultstable(searchObj, dgvOctopartResults)

    End Function

    Private Function resultstable(ByVal res As NexarAPIResponse, ByRef dgv As DataGridView)

        dgv.Columns.Clear()
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill

        Dim searchResultsList As New List(Of List(Of String))
        Dim searchResults As New DataTable
        Dim row As Integer = 0

        searchResults.Columns.Add("MPN", GetType(String))
        searchResults.Columns.Add("Short Description", GetType(String))
        searchResults.Columns.Add("Manufacturer", GetType(String))
        searchResults.Columns.Add("Category", GetType(String))
        searchResults.Columns.Add("Median Price 1000", GetType(String))
        searchResults.Columns.Add("No. of sellers", GetType(String))
        searchResults.Columns.Add("Image URL", GetType(String))


        Dim btn As New DataGridViewButtonColumn


        For Each result In res.data.supSearch.results
            searchResultsList.Add(New List(Of String))
            searchResultsList(row).Add(result.part.mpn)
            searchResultsList(row).Add(result.part.manufacturer.name)

            Dim shortDesc As String

            If result.part.shortDescription IsNot Nothing Then
                shortDesc = result.part.shortDescription
            Else
                shortDesc = "N/A"
            End If

            searchResultsList(row).Add(shortDesc)

            Dim ctg As String

            If result.part.category IsNot Nothing Then
                ctg = result.part.category.name
            Else
                ctg = "N/A"
            End If

            searchResultsList(row).Add(ctg)

            Dim mp1000 As String
            If result.part.medianPrice1000 IsNot Nothing Then
                mp1000 = result.part.medianPrice1000.convertedCurrency & Math.Round(result.part.medianPrice1000.convertedPrice, 3)
            Else
                mp1000 = "N/A"
            End If

            searchResultsList(row).Add(mp1000)
            searchResultsList(row).Add(result.part.sellers.Length())
            Dim imgURL As String
            If result.part.bestImage IsNot Nothing Then
                imgURL = result.part.bestImage.url
            Else
                imgURL = "N/A"
            End If
            searchResultsList(row).Add(imgURL)
            row = row + 1
        Next


        For Each r As List(Of String) In searchResultsList
            searchResults.Rows.Add(r(0), r(2), r(1), r(3), r(4), r(5), r(6))
        Next


        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        btn.HeaderText = "Sellers"
        btn.Text = "View Sellers"
        btn.Name = "btnSellers"
        btn.UseColumnTextForButtonValue = True


        searchResults.Columns.Add("Part Image", GetType(Image))

        For Each r As DataRow In searchResults.Rows
            Dim jpg1 As Image
            Dim imgreq As HttpWebRequest = WebRequest.Create(r("Image URL").ToString())
            imgreq.Method = "GET"
            Dim imgres As HttpWebResponse = imgreq.GetResponse()
            jpg1 = Image.FromStream(imgres.GetResponseStream())
            Dim jpg2 = New Bitmap(jpg1, 180, 180)
            jpg2.SetResolution(jpg1.Width, jpg1.Height)
            imgres.Close()
            r("Part Image") = jpg2
        Next
        searchResults.AcceptChanges()

        Dim img As New DataGridViewImageColumn

        img.DataPropertyName = "Part Image"
        img.Name = "Part Image"
        img.ImageLayout = DataGridViewImageCellLayout.Zoom


        dgv.Columns.Add(img)
        dgv.DataSource = searchResults
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.Columns.Add(btn)
        dgv.Columns("Image URL").Visible = False
        dgvOctopartResults.Enabled = True

        Me.Text = $"{searchResults.Rows.Count} parts found"

        enableControls()
        btnExportPR.Enabled = False

        If searchResults.Rows.Count = 0 Then
            MessageBox.Show($"No part matches found for {tbKeyword.Text}")
            If DialogResult.OK Then
                enableControls()
                Exit Function
            End If
        End If

        statusLabel.Enabled = True
        statusLabel.Text = "Results are grouped by their MPN. Click ""View Sellers"" to add a specific vendor offer to your PR! ( ͡° ͜ʖ ͡°)_/¯"
    End Function
    Private Sub buildPRdgv()

        dgvBuildPR.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Dim deleteBtn As New DataGridViewButtonColumn
        deleteBtn.HeaderText = "Remove Item"
        deleteBtn.Text = "Remove Item"
        deleteBtn.UseColumnTextForButtonValue = True

        Dim updateBtn As New DataGridViewButtonColumn
        updateBtn.HeaderText = "Update Item"
        updateBtn.Text = "Update Item"
        updateBtn.UseColumnTextForButtonValue = True

        dgvBuildPR.DataSource = prTable
        dgvBuildPR.Columns.Add(deleteBtn)
        dgvBuildPR.DefaultCellStyle.WrapMode = DataGridViewTriState.True

    End Sub
    Private Async Function populateSubcategories(ByVal category As String) As Task

        cbSubcategories.Items.Clear()

        Dim subcategoriesList As New List(Of String)
        subcategoriesList.Add("All Subcategories")

        For Each ctg In categoriesObj.data.supCategories
            If ctg.name.ToString = category Then
                If ctg.children.Length <> 0 Then
                    For Each subctg In ctg.children
                        subcategoriesList.Add(subctg.name)
                    Next
                Else
                    subcategoriesList.Clear()
                    subcategoriesList.Add("N/A")
                End If
            End If
        Next

        cbSubcategories.Items.AddRange(subcategoriesList.ToArray)
        cbSubcategories.SelectedIndex = 0

        If String.IsNullOrEmpty(cbSubcategories.Text) = False Then
            labelKeyword.Enabled = True
            tbKeyword.Enabled = True
            btnSearch.Enabled = True
        End If
        groupFilters.Enabled = True

    End Function

    Private Sub cbCategories_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cbCategories.SelectedIndexChanged
        populateSubcategories(cbCategories.Text)
    End Sub

    Private Async Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        selectedSource = cbSources.Text
        If selectedSource = "Octopart" Then
            Await searchOcto(tbKeyword.Text, cbCategories.Text, cbSubcategories.Text)
        ElseIf selectedSource = "Shopee" Then
            disableControls()
            Dim resultsDataTable = Await SearchShopee(tbKeyword.Text, dgvOctopartResults, ToolStripProgressBar1, statusLabel)
            dgvOctopartResults.DataSource = resultsDataTable
            dgvOctopartResults.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgvOctopartResults.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvOctopartResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            Dim btn As New DataGridViewButtonColumn()
            dgvOctopartResults.Columns.Add(btn)
            btn.HeaderText = "Item Links"
            btn.Text = "View on browser"
            btn.Name = "btnLink"
            btn.UseColumnTextForButtonValue = True

            Dim btnAddtoPR As New DataGridViewButtonColumn()
            dgvOctopartResults.Columns.Add(btnAddtoPR)
            btnAddtoPR.HeaderText = "Add to PR"
            btnAddtoPR.Text = "Add to Purchase Request"
            btnAddtoPR.Name = "btnAddToPR"
            btnAddtoPR.UseColumnTextForButtonValue = True
            'hide product page
            dgvOctopartResults.Columns("Product Page").Visible = False
            dgvOctopartResults.Columns("Image URL").Visible = False
            dgvOctopartResults.AllowUserToAddRows = False
            enableControls()
        Else
            MessageBox.Show("Please select a source!!!!!")
        End If
    End Sub

    Private Sub dgvOctopartResults_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOctopartResults.CellContentClick
        Dim sendergrid = DirectCast(sender, DataGridView)
        If selectedSource = "Octopart" Then
            If TypeOf sendergrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn Then
                populateSellersTable(sendergrid.CurrentRow.Cells(0).Value.ToString(), sendergrid.CurrentRow.Cells(1).Value.ToString(), sendergrid.CurrentRow.Cells(2).Value.ToString(), searchObj, viewSellers.dgvSellers)
                viewSellers.MPN = sendergrid.CurrentRow.Cells(0).Value.ToString()
                viewSellers.shortDesc = sendergrid.CurrentRow.Cells(1).Value.ToString()
                viewSellers.Manufacturer = sendergrid.CurrentRow.Cells(2).Value.ToString()
            End If
        ElseIf selectedSource = "Octopart" Then
            MessageBox.Show("No")
        End If
    End Sub

    Private Sub populateSellersTable(ByVal MPN As String, ByVal shortDesc As String, ByVal manufacturer As String, ByVal results As NexarAPIResponse, ByVal dgv As DataGridView)

        dgv.Columns.Clear()

        Dim sellersList As New List(Of List(Of String))
        Dim sellersTable As New DataTable
        Dim row As Integer = 0
        Dim companyLink As New DataGridViewButtonColumn
        Dim productLink As New DataGridViewButtonColumn
        Dim addToPR As New DataGridViewButtonColumn

        sellersTable.Columns.Add("Company Name", GetType(String))
        sellersTable.Columns.Add("Company Website", GetType(String))
        sellersTable.Columns.Add("Product Page", GetType(String))
        sellersTable.Columns.Add("Inventory Level", GetType(String))
        sellersTable.Columns.Add("Min. Order Qty.", GetType(String))
        sellersTable.Columns.Add("Prices", GetType(String))
        sellersTable.Columns.Add("Packaging", GetType(String))

        For Each o In results.data.supSearch.results
            If o.part.mpn = MPN Then
                For Each p In o.part.sellers
                    Dim companyName As String = p.company.name
                    Dim companyUrl As String = p.company.homepageUrl
                    For Each offer In p.offers
                        sellersList.Add(New List(Of String))
                        sellersList(row).Add(companyName)
                        sellersList(row).Add(companyUrl)
                        sellersList(row).Add(offer.clickUrl)
                        sellersList(row).Add(offer.inventoryLevel)
                        Dim moq As String
                        If offer.moq IsNot Nothing Then
                            moq = offer.moq
                        Else
                            moq = "N/A"
                        End If

                        Dim prices As New List(Of String)

                        If offer.prices.Length > 0 Then
                            Dim quantities As New List(Of Integer)
                            For Each price In offer.prices

                                Dim priceString As String = $"{price.convertedCurrency } {Math.Round(CDbl(price.convertedPrice), 3).ToString} || {price.quantity} pcs. min."
                                quantities.Add(price.quantity)
                                prices.Add(priceString)
                            Next
                            If IsNumeric(moq) Then
                                If quantities.Min > moq Then
                                    moq = quantities.Min
                                End If
                            End If
                            sellersList(row).Add(moq)
                            sellersList(row).Add(String.Join(Environment.NewLine, prices.ToArray))
                        Else
                            sellersList(row).Add(moq)
                            sellersList(row).Add("N/A")
                        End If

                        Dim pkg As String

                        If offer.packaging IsNot Nothing Then
                            pkg = offer.packaging
                        Else
                            pkg = "N/A"
                        End If

                        sellersList(row).Add(pkg)
                        row = row + 1
                    Next
                Next
            End If
        Next

        For Each l As List(Of String) In sellersList
            sellersTable.Rows.Add(l(0), l(1), l(2), l(3), l(4), l(5), l(6))
        Next

        dgv.ReadOnly = False
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        companyLink.HeaderText = "Seller Website"
        productLink.HeaderText = "Product Page"
        addToPR.HeaderText = "Add item to PR"
        companyLink.Text = "Go to Seller Homepage"
        productLink.Text = "Go to Product Page"
        addToPR.Text = "Add to PR"
        companyLink.UseColumnTextForButtonValue = True
        productLink.UseColumnTextForButtonValue = True
        addToPR.UseColumnTextForButtonValue = True


        dgv.DataSource = sellersTable
        dgv.Columns("Prices").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.Columns.Add(productLink)
        dgv.Columns.Add(companyLink)
        dgv.Columns.Add("Quantity", "Quantity")
        dgv.Columns.Add(addToPR)

        For Each col As DataGridViewColumn In dgv.Columns
            If col.Name <> "Quantity" Then
                col.ReadOnly = True
            End If
        Next

        dgv.Columns("Quantity").ReadOnly = False

        dgv.Columns("Company Website").Visible = False
        dgv.Columns("Product Page").Visible = False

        viewSellers.Text = $"Showing available sellers for MPN: {MPN}"
        viewSellers.Show()
        viewSellers.ToolStripStatusLabel1.Text = "(ノ ゜-゜)ノ"

    End Sub

    Private Sub btnExportPR_Click(sender As Object, e As EventArgs) Handles btnExportPR.Click

        exportPR()

    End Sub

    Private Sub exportPR()
        Dim prTableforExcel As New DataTable

        prTableforExcel.Columns.Add("MPN + Description", GetType(String))
        prTableforExcel.Columns.Add("Manufacturer", GetType(String))
        prTableforExcel.Columns.Add("Distributor", GetType(String))
        prTableforExcel.Columns.Add("Product Page", GetType(String))
        prTableforExcel.Columns.Add("Quantity", GetType(Integer))
        prTableforExcel.Columns.Add("Unit Price", GetType(String))
        prTableforExcel.Columns.Add("Total Price", GetType(String))

        For Each row As DataRow In prTable.Rows
            Dim mpnDesc As New List(Of String)
            mpnDesc.Add(row.Item(0).ToString)
            mpnDesc.Add(row.Item(1).ToString)
            Dim mpnDesc37 As IEnumerable(Of String) = SplitInParts(String.Join(" | ", mpnDesc.ToArray), 37)
            'For i As Integer = 0 To String.Join(" | ", mpnDesc.ToArray).Length - 1 Step 37
            '    mpnDesc37.Add(String.Join(" | ", mpnDesc.ToArray).Substring(i, 37))
            'Next

            If mpnDesc37.Count = 1 Then
                If row.Item(8).ToString.Contains("PHP") Then
                    prTableforExcel.Rows.Add(mpnDesc37(0), row.Item(2).ToString, row.Item(3).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString)
                Else
                    Dim unitPrice As String = $"PHP {row.Item(8).ToString}"
                    prTableforExcel.Rows.Add(mpnDesc37(0), row.Item(2).ToString, row.Item(3).ToString, row.Item(6).ToString, row.Item(7).ToString, unitPrice, row.Item(9).ToString)
                End If

            Else
                If row.Item(8).ToString.Contains("PHP") Then
                    prTableforExcel.Rows.Add(mpnDesc37(0), row.Item(2).ToString, row.Item(3).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString)
                    For i As Integer = 1 To mpnDesc37.Count - 1
                        prTableforExcel.Rows.Add(mpnDesc37(i), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Next
                Else
                    Dim unitPrice As String = $"PHP {row.Item(8).ToString}"
                    prTableforExcel.Rows.Add(mpnDesc37(0), row.Item(2).ToString, row.Item(3).ToString, row.Item(6).ToString, row.Item(7).ToString, unitPrice, row.Item(9).ToString)
                    For i As Integer = 1 To mpnDesc37.Count - 1
                        prTableforExcel.Rows.Add(mpnDesc37(i), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Next
                End If

            End If

        Next

        dlgSaveFile.Filter = "Excel (*.xlsx)|*.xlsx"
        dlgSaveFile.FilterIndex = 2
        dlgSaveFile.RestoreDirectory = True
        If dlgSaveFile.ShowDialog() = DialogResult.OK Then
            addtoExcel(prTableforExcel, dlgSaveFile.FileName)
        End If
    End Sub

    Private Sub exportPR_keyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If btnExportPR.Enabled = True Then
            If e.Control AndAlso (e.KeyCode = 83) Then
                exportPR()
            End If
        End If
    End Sub

    Public Function SplitInParts(s As String, partLength As Integer) As IEnumerable(Of String)
        If String.IsNullOrEmpty(s) Then
            Throw New ArgumentNullException("String cannot be null or empty.")
        End If
        If partLength <= 0 Then
            Throw New ArgumentException("Split length has to be positive.")
        End If
        Return Enumerable.Range(0, Math.Ceiling(s.Length / partLength)).Select(Function(i) s.Substring(i * partLength, If(s.Length - (i * partLength) >= partLength, partLength, Math.Abs(s.Length - (i * partLength)))))
    End Function

    Private Sub dgvBuildPR_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBuildPR.CellContentClick
        Dim sendergrid = DirectCast(sender, DataGridView)
        If TypeOf sendergrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn Then
            If e.ColumnIndex = 10 Then

                Dim i As Integer = e.RowIndex
                Dim row As DataGridViewRow

                For Each r As DataGridViewRow In dgvBuildPR.Rows
                    If r.Index = i Then
                        r.Selected = True
                    End If
                Next

                row = dgvBuildPR.SelectedRows(0)

                If row.IsNewRow = False Then
                    DirectCast(row.DataBoundItem, DataRowView).Delete()
                End If

                Dim totalPriceList As New List(Of Double)
                Dim totalQtyList As New List(Of Integer)

                For Each r As DataRow In prTable.Rows
                    If r.Item(9).ToString.Contains("PHP") Then
                        totalPriceList.Add(r.Item(9).ToString.Split(" ")(1))
                    End If
                    totalQtyList.Add(r.Item(7))
                Next

                Dim prTotalCost As String = totalPriceList.Sum.ToString
                Dim prTotalCount As String = prTable.Rows.Count

                If prTotalCount > 0 Then
                    labelPRStat.Text = $"[ Total Items: {prTotalCount}, Total Cost: {prTotalCost}, Total Order Qty: {totalQtyList.Sum} ]"
                Else
                    labelPRStat.Text = ""
                End If

            End If
        End If
    End Sub

    Private Sub dgvBuildPR_CellContentChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBuildPR.CellValueChanged

        Dim totalPrice As Double
        Dim unitPrice As Double
        Dim newUnitPrice As String
        Dim newTotalPrice As String

        If e.ColumnIndex = 7 Then
            If IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString) Then
                If dgvBuildPR.Rows(e.RowIndex).Cells(9).Value.ToString.Contains("PHP") Then
                    totalPrice = CDbl(dgvBuildPR.Rows(e.RowIndex).Cells(9).Value.ToString.Split(" ")(1))
                Else
                    totalPrice = CDbl(dgvBuildPR.Rows(e.RowIndex).Cells(9).Value)
                End If

                If dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString.Contains("PHP") Then
                    unitPrice = CDbl(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString.Split(" ")(1))
                Else
                    unitPrice = CDbl(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value)
                End If

                Try
                    If IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(7).Value.ToString) Then
                        If dgvBuildPR.Rows(e.RowIndex).Cells(7).Value.ToString >= CInt(dgvBuildPR.Rows(e.RowIndex).Cells(4).Value) Then
                            newTotalPrice = $"PHP {unitPrice * dgvBuildPR.Rows(e.RowIndex).Cells(7).Value}"
                            dgvBuildPR.Rows(e.RowIndex).Cells(9).Value = newTotalPrice
                        Else
                            MessageBox.Show($"Please enter a value equal to or above MOQ: {dgvBuildPR.Rows(e.RowIndex).Cells(4).Value.ToString}")
                            dgvBuildPR.Rows(e.RowIndex).Cells(7).Value = dgvBuildPR.Rows(e.RowIndex).Cells(4).Value
                        End If

                    Else
                        MessageBox.Show("Please enter a valid numeric value.")
                        dgvBuildPR.Rows(e.RowIndex).Cells(7).Value = prTable.Rows(e.RowIndex).Item(7)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Please enter a valid numeric value.")
                    dgvBuildPR.Rows(e.RowIndex).Cells(7).Value = prTable.Rows(e.RowIndex).Item(7)
                End Try
            ElseIf IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString.Split(" ")(1)) Then
                totalPrice = CDbl(dgvBuildPR.Rows(e.RowIndex).Cells(9).Value.ToString.Split(" ")(1))
                unitPrice = CDbl(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString.Split(" ")(1))
                Try
                    If IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(7).Value.ToString) Then
                        If dgvBuildPR.Rows(e.RowIndex).Cells(7).Value.ToString >= CInt(dgvBuildPR.Rows(e.RowIndex).Cells(4).Value) Then
                            newTotalPrice = $"PHP {unitPrice * dgvBuildPR.Rows(e.RowIndex).Cells(7).Value}"
                            dgvBuildPR.Rows(e.RowIndex).Cells(9).Value = newTotalPrice
                        Else
                            MessageBox.Show($"Please enter a value equal to or above MOQ: {dgvBuildPR.Rows(e.RowIndex).Cells(4).Value.ToString}")
                            dgvBuildPR.Rows(e.RowIndex).Cells(7).Value = dgvBuildPR.Rows(e.RowIndex).Cells(4).Value
                        End If

                    Else
                        MessageBox.Show("Please enter a valid numeric value.")
                        dgvBuildPR.Rows(e.RowIndex).Cells(7).Value = prTable.Rows(e.RowIndex).Item(7)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Please enter a valid numeric value.")
                    dgvBuildPR.Rows(e.RowIndex).Cells(7).Value = prTable.Rows(e.RowIndex).Item(7)
                End Try
            Else
                MessageBox.Show("Please check product page for price and update Unit Price column manually")
                dgvBuildPR.Rows(e.RowIndex).Cells(8).ReadOnly = False
                dgvBuildPR.Rows(e.RowIndex).Cells(8).Style.Font = New Font(FontFamily.GenericSansSerif, 8.75, FontStyle.Bold)
                dgvBuildPR.Rows(e.RowIndex).Cells(8).Style.BackColor = Drawing.Color.FromArgb(228, 229, 224)
            End If
        ElseIf e.ColumnIndex = 8 Then
            Try

                If IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString) Or IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(8).Value.ToString.Split(" ")(1)) Then
                    newUnitPrice = $"PHP {dgvBuildPR.Rows(e.RowIndex).Cells(8).Value}"
                    dgvBuildPR.Rows(e.RowIndex).Cells(8).Value = newUnitPrice
                    newTotalPrice = $"PHP {newUnitPrice.ToString.Split(" ")(1)}"
                    dgvBuildPR.Rows(e.RowIndex).Cells(9).Value = newTotalPrice
                Else
                    MessageBox.Show("Please enter a valid numeric value")
                End If

            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub dgvBuildPR_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvBuildPR.CellValidating

        If e.ColumnIndex = 7 Then
            dgvBuildPR.Rows(e.RowIndex).ErrorText = ""
            Dim newQty As Integer
            If IsNumeric(e.FormattedValue) = True Then
                If IsNumeric(dgvBuildPR.Rows(e.RowIndex).Cells(4).Value) Then
                    If e.FormattedValue > 0 Then
                        If e.FormattedValue < dgvBuildPR.Rows(e.RowIndex).Cells(4).Value Then
                            e.Cancel = True
                            dgvBuildPR.Rows(e.RowIndex).ErrorText = $"Please enter a value equal to or above {dgvBuildPR.Rows(e.RowIndex).Cells(4).Value}"
                            MessageBox.Show($"Please enter a value equal to or above {dgvBuildPR.Rows(e.RowIndex).Cells(4).Value}")
                        End If
                    Else
                        e.Cancel = True
                        dgvBuildPR.Rows(e.RowIndex).ErrorText = "Please enter a valid numeric value"
                        MessageBox.Show("Please enter a valid numeric value")
                    End If

                ElseIf e.FormattedValue <= 0 Then
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
                dgvBuildPR.Rows(e.RowIndex).ErrorText = "Please enter a valid numeric value"
                MessageBox.Show("Please enter a valid numeric value")
            End If
        ElseIf e.ColumnIndex = 8 Then
            dgvBuildPR.Rows(e.RowIndex).ErrorText = ""
            If IsNumeric(e.FormattedValue) = False Then
                Try
                    If IsNumeric(e.FormattedValue.ToString.Split(" ")(1)) = False Then
                        e.Cancel = True
                        MessageBox.Show("Please enter a valid numeric value")
                    Else
                        dgvBuildPR.Rows(e.RowIndex).Cells(9).Value = $"PHP {e.FormattedValue.ToString.Split(" ")(1) * dgvBuildPR.Rows(e.RowIndex).Cells(7).Value}"
                    End If
                Catch ex As Exception
                    MessageBox.Show($"Please recheck your input for {dgvBuildPR.Columns(e.ColumnIndex).Name}", "Error reading input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            ElseIf IsNumeric(e.FormattedValue) Then
                dgvBuildPR.Rows(e.RowIndex).Cells(9).Value = $"PHP {e.FormattedValue * dgvBuildPR.Rows(e.RowIndex).Cells(7).Value}"
            ElseIf IsNumeric(e.FormattedValue.ToString.Split(" ")(1)) Then
                dgvBuildPR.Rows(e.RowIndex).Cells(9).Value = $"PHP {e.FormattedValue.ToString.Split(" ")(1) * dgvBuildPR.Rows(e.RowIndex).Cells(7).Value}"
            End If
        End If

    End Sub

    Private Sub addtoExcel(ByVal dt As DataTable, ByVal filename As String)

        Dim xlApp As New Microsoft.Office.Interop.Excel.Application()
        Dim xlNewSheet
        If xlApp Is Nothing Then
            MessageBox.Show("Excel was not found in your PC.")
            Return
        End If

        xlApp.DisplayAlerts = False
        Dim filepath As String = "C:\Users\PC\Documents\FORM-001_Purchase Request Form"
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(filepath, 0, False, 5, "", "", False, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", True, False, 0, True, False, False)
        Dim worksheets As Excel.Sheets = xlWorkBook.Worksheets

        Try
            xlNewSheet = DirectCast(worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing), Excel.Worksheet)
            xlNewSheet.Name = "csvSheet"
            'xlNewSheet.Cells(1, 1) = "New Sheet Content"

            For k = 0 To dt.Columns.Count - 1
                xlNewSheet.Cells(1, k + 1) = dt.Columns.Item(k).ToString
            Next

            For i = 0 To dt.Rows.Count - 1
                For j = 0 To dt.Columns.Count - 1
                    xlNewSheet.Cells(i + 2, j + 1) = dt.Rows(i).Item(j).ToString
                Next
            Next

            xlNewSheet = xlWorkBook.Sheets("PR Form (2)")
            Dim wsFormatted As Excel.Worksheet = CType(xlWorkBook.Sheets("PR Form (2)"), Excel.Worksheet)

            Select Case dt.Rows.Count
                Case <= 10
                    wsFormatted.ExportAsFixedFormat(0, $"{filename.Split(".")(0)}.pdf",,,, 1, 1)
                Case 10 <= 25
                    wsFormatted.ExportAsFixedFormat(0, $"{filename.Split(".")(0)}.pdf",,,, 1, 2)
                Case 25 <= 40
                    wsFormatted.ExportAsFixedFormat(0, $"{filename.Split(".")(0)}.pdf",,,, 1, 3)
            End Select
            'xlNewSheet.Select()
            xlWorkBook.SaveAs(filename)
            xlWorkBook.Close()
            releaseObject(xlNewSheet)
            releaseObject(worksheets)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
            MessageBox.Show($"File: [{filename}] succesfully saved.")
            Exit Sub
        Catch ex As Exception
            Dim sheet1 As Excel.Worksheet = CType(xlWorkBook.Sheets("Sheet1"), Excel.Worksheet)
            sheet1.Delete()
            Dim ws As Excel.Worksheet = CType(xlWorkBook.Sheets("csvSheet"), Excel.Worksheet)
            ws.Select()
            ws.Range("A1:Z100").Select()
            ws.Range("A1:Z100").ClearContents()
            xlWorkBook.Sheets("PR Form (2)").Select()
            For k = 0 To dt.Columns.Count - 1
                ws.Cells(1, k + 1) = dt.Columns.Item(k).ToString
            Next

            For i = 0 To dt.Rows.Count - 1
                For j = 0 To dt.Columns.Count - 1
                    ws.Cells(i + 2, j + 1) = dt.Rows(i).Item(j).ToString
                Next
            Next
            xlWorkBook.SaveAs(filename)
            Dim wsFormatted As Excel.Worksheet = CType(xlWorkBook.Sheets("PR Form (2)"), Excel.Worksheet)

            Select Case dt.Rows.Count
                Case <= 10
                    wsFormatted.ExportAsFixedFormat(0, $"{filename.Split(".")(0)}.pdf",,,, 1, 1)
                Case <= 25
                    wsFormatted.ExportAsFixedFormat(0, $"{filename.Split(".")(0)}.pdf",,,, 1, 2)
                Case <= 40
                    wsFormatted.ExportAsFixedFormat(0, $"{filename.Split(".")(0)}.pdf",,,, 1, 3)
            End Select

            xlWorkBook.Close()
            releaseObject(xlNewSheet)
            releaseObject(worksheets)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
            MessageBox.Show($"File: [{filename}] succesfully saved.")
            Exit Sub
        End Try

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub cbSources_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSources.SelectedIndexChanged
        If cbSources.SelectedItem = "Octopart" Then
            prTable.Columns.Clear()
            dgvBuildPR.Columns.Clear()
            groupFilters.Visible = True
            btnSearch.Location = New Point(72, 296)
            statusLabel.Text = "Currently fetching data from authorized distributors... ノ( ゜-゜ノ)"
            If cbCategories.Items.Count = 0 Then
                disableControls()
                populateCategories()
            End If
            prTable.Columns.Add("MPN", GetType(String))
            prTable.Columns.Add("Short Description", GetType(String))
            prTable.Columns.Add("Manufacturer", GetType(String))
            prTable.Columns.Add("Distributor", GetType(String))
            prTable.Columns.Add("MOQ", GetType(String))
            prTable.Columns.Add("Prices", GetType(String))
            prTable.Columns.Add("Product Page", GetType(String))
            prTable.Columns.Add("Quantity", GetType(Integer))
            prTable.Columns.Add("Unit Price", GetType(String))
            prTable.Columns.Add("Total Price", GetType(String))
            buildPRdgv()
        ElseIf cbSources.SelectedItem = "Shopee" Then
            groupFilters.Visible = False
            btnSearch.Enabled = True
            btnSearch.Location = New Point(72, 144)
        End If
    End Sub
End Class