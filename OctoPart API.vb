Imports RestSharp
Imports Newtonsoft.Json
Imports System.Net
Imports System.Linq
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop

Public Class OctoPart_API

    Public categoriesObj
    Public prTable As New DataTable
    Dim searchObj
    Dim csvSavepath As String

    'Public Async Function sendAPIRequest(ByVal query As String) As Task(Of Object)

    '    Dim options = New RestClientOptions("http://octopart.com/api/v3")
    '    options.ThrowOnAnyError = True
    '    options.MaxTimeout = queryTimeOut
    '    Dim client = New RestClient(options)
    '    Dim request = New RestRequest()
    '    request.Timeout = queryTimeOut
    '    request.Method = Method.Post
    '    request.AddHeader("Authorization", token)
    '    request.AddQueryParameter("query", query)
    '    Dim response = New RestResponse
    '    response = Await client.PostAsync(request)
    '    'If response.IsSuccessStatusCode = True Then
    '    '    Return response.Content
    '    'Else
    '    '    Return False
    '    'End If
    '    If response.IsSuccessStatusCode Then
    '        'response has a statuscode of 200
    '        'but it might have a parse error, which still is status 200.
    '        If response.Content.Contains("error") Or response.Content.Contains("error_message") Or response.Content.Contains("errors") Then
    '            'response has a status code 200, but has a monday.com error.
    '            Return {"error", response.Content}
    '        Else
    '            'response has a status code 200, with readable results.
    '            Return {"success", response.Content}
    '        End If
    '    Else
    '        Throw New System.Exception("An error has occured at function: SendMondayRequestVersion2")
    '    End If

    'End Function

    Private Sub disableControls()
        For Each c As Control In Me.Controls
            c.Enabled = False
        Next
    End Sub

    Private Sub enableControls()
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next
    End Sub

    Private Sub OctoPart_API_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        statusLabel.Text = "Currently fetching data from authorized distributors... ノ( ゜-゜ノ)"
        disableControls()
        populateCategories()
        prTable.Columns.Add("MPN", GetType(String))
        prTable.Columns.Add("Short Description", GetType(String))
        prTable.Columns.Add("Manufacturer", GetType(String))
        prTable.Columns.Add("Distributor", GetType(String))
        prTable.Columns.Add("Prices", GetType(String))
        prTable.Columns.Add("Product Page", GetType(String))
        prTable.Columns.Add("Quantity", GetType(Integer))
        prTable.Columns.Add("Unit Price", GetType(String))
        prTable.Columns.Add("Total Price", GetType(String))
        buildPRdgv()
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


        cbCategories.Items.AddRange(categoriesList.ToArray)
        cbCategories.SelectedIndex = 0

        groupFilters.Enabled = True
        cbCategories.Enabled = True
        labelCategories.Enabled = True

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
        btnGetTotal.Enabled = False
        btnExportPR.Enabled = False

        If searchResults.Rows.Count = 0 Then
            MessageBox.Show($"No part matches found for {tbKeyword.Text}")
        End If

        statusLabel.Enabled = True
        statusLabel.Text = "Results are grouped by their MPN. Click ""View Sellers"" to add a specific vendor offer to your PR! ( ͡° ͜ʖ ͡°)_/¯"
    End Function
    Private Sub buildPRdgv()

        dgvBuildPR.ReadOnly = False

        For Each col As DataGridViewColumn In dgvBuildPR.Columns
            If col.Name <> "Quantity" Then
                col.ReadOnly = True
            End If
        Next

        dgvBuildPR.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Dim updateBtn As New DataGridViewButtonColumn
        updateBtn.HeaderText = "Remove Item"
        updateBtn.Text = "Remove Item"
        updateBtn.UseColumnTextForButtonValue = True
        dgvBuildPR.DataSource = prTable

        dgvBuildPR.Columns.Add(updateBtn)
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

    End Function

    Private Sub cbCategories_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cbCategories.SelectedIndexChanged
        populateSubcategories(cbCategories.Text)
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        searchOcto(tbKeyword.Text, cbCategories.Text, cbSubcategories.Text)
    End Sub

    Private Sub dgvOctopartResults_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOctopartResults.CellContentClick

        Dim sendergrid = DirectCast(sender, DataGridView)

        If TypeOf sendergrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn Then
            populateSellersTable(sendergrid.CurrentRow.Cells(0).Value.ToString(), sendergrid.CurrentRow.Cells(1).Value.ToString(), sendergrid.CurrentRow.Cells(2).Value.ToString(), searchObj, viewSellers.dgvSellers)
            viewSellers.MPN = sendergrid.CurrentRow.Cells(0).Value.ToString()
            viewSellers.shortDesc = sendergrid.CurrentRow.Cells(1).Value.ToString()
            viewSellers.Manufacturer = sendergrid.CurrentRow.Cells(2).Value.ToString()
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

    Private Sub btnGetTotal_Click(sender As Object, e As EventArgs) Handles btnGetTotal.Click
        'Dim totals As New List(Of Double)

        'For i As Integer = 0 To prTable.Rows.Count
        '    If IsNumeric(prTable.Rows(i).Item(8).ToString.Split(" ")(1)) Then
        '        Dim total As Double = prTable.Rows(i).Item(8).ToString.Split(" ")(1)
        '        totals.Add(total)
        '    End If

        'Next

        'prTable.Rows.Add(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, $"PHP {totals.Sum}")

        MessageBox.Show("Code under construction", "¯\( ◡ ‿ ◡)/¯")
    End Sub

    Private Function dtTableToCSV(dt As DataTable, ByVal filename As String, Optional headers As Boolean = True, Optional delim As String = ",")
        Try
            Dim txt As String
            Dim csvWriter As IO.StreamWriter = IO.File.AppendText(filename)
            Dim fileloc As String = filename
            Dim n = 0
            If IO.File.Exists(filename) Then
                txt += vbCrLf
                txt += "PR generated on " + DateTime.Now().ToString("g")
                txt += vbCrLf
                If headers = True Then
                    For Each column As DataColumn In dt.Columns
                        If n = 0 Then
                            txt += column.ColumnName.Replace(",", "--comma--")
                        Else
                            txt += delim + column.ColumnName.Replace(",", "--comma--")
                        End If
                        n += 1
                    Next
                End If
                txt += vbCrLf
                n = 0
                For Each row As DataRow In dt.Rows
                    Dim line As String = ""
                    For Each column As DataColumn In dt.Columns
                        line += delim & row(column.ColumnName).ToString().Replace(",", "--comma--")
                    Next
                    If dt.Rows.Count - 1 = n Then
                        txt += line.Substring(1)
                    Else
                        txt += line.Substring(1) & vbCrLf
                    End If
                    n += 1
                Next
            End If

            csvWriter.Write(txt)
            csvWriter.Close()



        Catch ex As Exception
            MessageBox.Show("Please select a valid save directory", "Invalid save location.", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub btnExportPR_Click(sender As Object, e As EventArgs) Handles btnExportPR.Click

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
                prTableforExcel.Rows.Add(mpnDesc37(0), row.Item(2).ToString, row.Item(3).ToString, row.Item(5).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString)
            Else
                prTableforExcel.Rows.Add(mpnDesc37(0), row.Item(2).ToString, row.Item(3).ToString, row.Item(5).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString)
                For i As Integer = 1 To mpnDesc37.Count - 1
                    prTableforExcel.Rows.Add(mpnDesc37(i), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                Next
            End If

        Next

        dlgSaveFile.Filter = "Excel (*.xlsx)|*.xlsx"
        dlgSaveFile.FilterIndex = 2
        dlgSaveFile.RestoreDirectory = True
        If dlgSaveFile.ShowDialog() = DialogResult.OK Then
            'dtTableToCSV(prTableforExcel, dlgSaveFile.FileName)
            addtoExcel(prTableforExcel, dlgSaveFile.FileName)
        End If

        'If csvSavepath <> "" Or csvSavepath IsNot Nothing Or csvSavepath.Length >= 3 Then
        '    Dim filename As String = csvSavepath + "\" + DateTime.Now.ToString("yyyyMMdd_HHmmss_") + ("YTD.csv")
        '    dtTableToCSV(prTableforCSV, filename)
        'Else
        '    MessageBox.Show("Please select a folder to save your file in.", "No folder selected", MessageBoxButtons.OK)
        'End If
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
            If e.ColumnIndex = 9 Then

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
            xlNewSheet.Select()
            xlWorkBook.SaveAs(filename)
            xlWorkBook.Close()
            releaseObject(xlNewSheet)
            releaseObject(worksheets)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
        Catch ex As Exception
            Dim sheet1 As Excel.Worksheet = CType(xlWorkBook.Sheets("Sheet1"), Excel.Worksheet)
            sheet1.Delete()
            Dim ws As Excel.Worksheet = CType(xlWorkBook.Sheets("csvSheet"), Excel.Worksheet)
            ws.Select()
            ws.Range("A1:Z100").Select()
            ws.Range("A1:Z100").ClearContents()
            For k = 0 To dt.Columns.Count - 1
                ws.Cells(1, k + 1) = dt.Columns.Item(k).ToString
            Next

            For i = 0 To dt.Rows.Count - 1
                For j = 0 To dt.Columns.Count - 1
                    ws.Cells(i + 2, j + 1) = dt.Rows(i).Item(j).ToString
                Next
            Next
            xlWorkBook.SaveAs(filename)
            xlWorkBook.Close()
            releaseObject(xlNewSheet)
            releaseObject(worksheets)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
            Exit Sub
        End Try

        MessageBox.Show($"File: [{filename}] succesfully saved.")

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

End Class