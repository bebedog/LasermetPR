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

    Private Sub buildprTable(ByVal _mpn As String, ByVal desc As String, ByVal mfr As String, ByVal seller As String, ByVal productURL As String, ByVal qty As Integer, ByVal dgv As DataGridView, ByVal prices As String)

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
            .Item(4) = prices
            .Item(5) = productURL
            .Item(6) = qty
            If prices = "Please check product page" Then
                .Item(7) = "Price information not found. Please check product page"
            Else
                .Item(7) = $"{unitPriceCurr} {Math.Round(unitPrice, 3)}"
            End If
            If prices = "Please check product page" Then
                .Item(8) = .Item(7)
            Else
                .Item(8) = $"{unitPriceCurr} {Math.Round(unitPrice * qty, 2)}"
            End If


        End With

        OctoPart_API.prTable.Rows.InsertAt(newRow, 1)
        OctoPart_API.dgvBuildPR.Enabled = True
        OctoPart_API.btnGetTotal.Enabled = True
        OctoPart_API.btnExportPR.Enabled = True
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
                    If IsNumeric(sendergrid.CurrentRow.Cells(2).Value.ToString) Then
                        If IsNumeric(sendergrid.CurrentRow.Cells(8).Value) Then
                            Dim moq As Integer = sendergrid.CurrentRow.Cells(8).Value
                            If moq <= sendergrid.CurrentRow.Cells(2).Value Then
                                qty = sendergrid.CurrentRow.Cells(2).Value
                                companyName = sendergrid.CurrentRow.Cells(4).Value
                                productLink = sendergrid.CurrentRow.Cells(6).Value
                                buildprTable(MPN, shortDesc, Manufacturer, companyName, productLink, qty, OctoPart_API.dgvBuildPR, sendergrid.CurrentRow.Cells(9).Value)
                                Me.Close()
                            Else
                                MessageBox.Show($"Please enter a quantity at least equal to {moq}")
                            End If
                        Else
                            qty = sendergrid.CurrentRow.Cells(2).Value
                            companyName = sendergrid.CurrentRow.Cells(4).Value
                            productLink = sendergrid.CurrentRow.Cells(6).Value
                            buildprTable(MPN, shortDesc, Manufacturer, companyName, productLink, qty, OctoPart_API.dgvBuildPR, sendergrid.CurrentRow.Cells(9).Value)
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