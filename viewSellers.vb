Imports Newtonsoft.Json

Public Class viewSellers

    Public MPN As String
    Public shortDesc As String
    Public Manufacturer As String
    Public companyName As String
    Public productLink As String
    Public qty As Integer

    Private Sub viewSellers_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub buildprTable(ByVal _mpn As String, ByVal desc As String, ByVal mfr As String, ByVal seller As String, ByVal productURL As String, ByVal qty As Integer, ByVal dgv As DataGridView, ByVal prices As String, ByVal moq As String)

        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells.AllCells
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim newRow As DataRow = OctoPart_API.prTable.NewRow()
        Dim unitPriceCurr As String
        Dim unitPrice As Double
        Dim row As Integer = 0

        Dim pricesArray As New List(Of List(Of String))

        If String.IsNullOrEmpty(prices) Or String.IsNullOrWhiteSpace(prices) Or prices = "N/A" Then
            prices = "Please check product page"
        Else
            For Each p In prices.Split(Environment.NewLine)
                Dim varPrice As String = p.Replace(" || ", "|")
                pricesArray.Add(New List(Of String))
                pricesArray(row).Add(varPrice.Split("|")(0).Trim)
                pricesArray(row).Add(varPrice.Split("|")(1).Replace(" pcs. min.", "").Trim)
                row = row + 1
            Next
        End If

        For Each l As List(Of String) In pricesArray
            If qty >= l(1) Then
                unitPrice = l(0).Split(" ")(1)
                unitPriceCurr = l(0).Split(" ")(0)
            End If
        Next



        With newRow
            .Item(0) = _mpn
            .Item(1) = desc
            .Item(2) = mfr
            .Item(3) = seller
            .Item(4) = moq
            .Item(5) = prices
            .Item(6) = productURL
            .Item(7) = qty
            If prices = "Please check product page" Then
                .Item(8) = "Price information not found. Please check product page"
            Else
                .Item(8) = $"{unitPriceCurr} {Math.Round(unitPrice, 3)}"
            End If
            If prices = "Please check product page" Then
                .Item(9) = .Item(7)
            Else
                .Item(9) = $"{unitPriceCurr} {Math.Round(unitPrice * qty, 2)}"
            End If


        End With

        OctoPart_API.prTable.Rows.InsertAt(newRow, 1)
        OctoPart_API.dgvBuildPR.Enabled = True
        OctoPart_API.btnExportPR.Enabled = True
        OctoPart_API.dgvBuildPR.ReadOnly = False

        For Each c As DataGridViewColumn In OctoPart_API.dgvBuildPR.Columns
            If c.Name <> "Quantity" Then
                c.ReadOnly = True
            End If

            OctoPart_API.dgvBuildPR.Columns("Quantity").ReadOnly = False

        Next

        OctoPart_API.dgvBuildPR.EnableHeadersVisualStyles = False

        OctoPart_API.dgvBuildPR.Columns("Quantity").HeaderCell.Style.Font = New Font(FontFamily.GenericSansSerif, 8.75, FontStyle.Bold)
        OctoPart_API.dgvBuildPR.Columns("Quantity").HeaderCell.Style.BackColor = Drawing.Color.FromArgb(228, 229, 224)

        OctoPart_API.dgvBuildPR.Columns("Quantity").DefaultCellStyle.Font = New Font(FontFamily.GenericSansSerif, 8.75, FontStyle.Bold)
        OctoPart_API.dgvBuildPR.Columns("Quantity").DefaultCellStyle.BackColor = Drawing.Color.FromArgb(228, 229, 224)

    End Sub

    Private Sub dgvSellers_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSellers.CellContentClick

        Dim sendergrid = DirectCast(sender, DataGridView)
        Dim url As String
        If TypeOf sendergrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn Then

            If e.ColumnIndex = 1 Then
                url = sendergrid.CurrentRow.Cells(5).Value.ToString
                Dim openURL As New ProcessStartInfo()
                openURL.Arguments = url
                openURL.UseShellExecute = True
                openURL.FileName = "msedge.exe"
                Process.Start(openURL)
            ElseIf e.ColumnIndex = 0 Then
                url = sendergrid.CurrentRow.Cells(6).Value.ToString
                Dim openURL As New ProcessStartInfo()
                openURL.Arguments = url
                openURL.UseShellExecute = True
                openURL.FileName = "msedge.exe"
                Process.Start(openURL)
            ElseIf e.ColumnIndex = 3 Then
                If sendergrid.CurrentRow.Cells(2).Value IsNot Nothing Then
                    Dim moq As String
                    If IsNumeric(sendergrid.CurrentRow.Cells(2).Value.ToString) Then
                        If IsNumeric(sendergrid.CurrentRow.Cells(8).Value) Then
                            moq = sendergrid.CurrentRow.Cells(8).Value
                            qty = sendergrid.CurrentRow.Cells(2).Value
                            companyName = sendergrid.CurrentRow.Cells(4).Value
                            productLink = sendergrid.CurrentRow.Cells(6).Value
                            If moq <= sendergrid.CurrentRow.Cells(2).Value Then
                                buildprTable(MPN, shortDesc, Manufacturer, companyName, productLink, qty, OctoPart_API.dgvBuildPR, sendergrid.CurrentRow.Cells(9).Value, moq)

                                Dim totalPriceList As New List(Of Double)
                                Dim totalQtyList As New List(Of Integer)

                                For Each r As DataRow In OctoPart_API.prTable.Rows
                                    If r.Item(9).ToString.Contains("PHP") Then
                                        totalPriceList.Add(r.Item(9).ToString.Split(" ")(1))
                                    End If
                                    totalQtyList.Add(r.Item(7))
                                Next

                                Dim prTotalCost As String = totalPriceList.Sum.ToString
                                Dim prTotalCount As String = OctoPart_API.prTable.Rows.Count

                                OctoPart_API.labelPRStat.Text = $"[ Total Items: {prTotalCount}, Total Cost: {prTotalCost}, Total Order Qty: {totalQtyList.Sum} ]"
                                Me.Close()
                            Else
                                MessageBox.Show($"{companyName} only accepts orders of {moq} pcs. and above for this item.")
                            End If
                        Else
                            moq = "N/A"
                            qty = sendergrid.CurrentRow.Cells(2).Value
                            companyName = sendergrid.CurrentRow.Cells(4).Value
                            productLink = sendergrid.CurrentRow.Cells(6).Value
                            buildprTable(MPN, shortDesc, Manufacturer, companyName, productLink, qty, OctoPart_API.dgvBuildPR, sendergrid.CurrentRow.Cells(9).Value, moq)
                            Dim totalPriceList As New List(Of Double)
                            Dim totalQtyList As New List(Of Integer)

                            For Each r As DataRow In OctoPart_API.prTable.Rows
                                If r.Item(9).ToString.Contains("PHP") Then
                                    totalPriceList.Add(r.Item(9).ToString.Split(" ")(1))
                                End If
                                totalQtyList.Add(r.Item(7))
                            Next

                            Dim prTotalCost As String = totalPriceList.Sum.ToString
                            Dim prTotalCount As String = OctoPart_API.prTable.Rows.Count

                            OctoPart_API.labelPRStat.Text = $"[ Total Items: {prTotalCount}, Total Cost: {prTotalCost}, Total Order Qty: {totalQtyList.Sum} ]"

                        End If
                    Else
                        MessageBox.Show("Please enter a valid quantity")
                    End If
                Else
                    MessageBox.Show("Please enter a valid quantity")
                End If

            End If

            End If
    End Sub

    Private Sub btnGoBack_Click(sender As Object, e As EventArgs) Handles btnGoBack.Click
        Me.Close()
    End Sub
End Class