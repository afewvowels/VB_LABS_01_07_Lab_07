'Program:		Lab #7
'Programmer:	Keith Smith
'Date:			31 October 2018
'Description:	Calculate sales price using the BookSale class.
'				Instantiate TheBookSale as a new object of the BookSale class.

Option Explicit On
Option Strict On

Public Class SalesForm
    ' Declare the new object.
    Dim TheBookSale As BookSale


    Private Sub CalculateSaleToolStripMenuItem_Click(ByVal sender As System.Object,
     ByVal e As System.EventArgs) Handles CalculateSaleToolStripMenuItem.Click
        ' Calculate the extended price for the sale.
        Dim TempQuantityInteger As Integer
        Dim TempPriceDecimal As Decimal

        Try
            ' Parse text fields to temporary variables
            TempQuantityInteger = Convert.ToInt32(QuantityTextBox.Text)
            TempPriceDecimal = Convert.ToDecimal(PriceTextBox.Text)

            ' Instantiate new object with overloaded constructor
            TheBookSale = New BookSale(TitleTextBox.Text, TempQuantityInteger, TempPriceDecimal)

            ' Display the calculated extended price result
            ExtendedPriceTextBox.Text = TheBookSale.ExtendedPrice.ToString("c")
        Catch ex As FormatException
            MessageBox.Show("Must use numeric values for quantity and price",
                            "Format Exception",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
        Catch ex As ArgumentOutOfRangeException
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        ' Clear the screen controls.
        QuantityTextBox.Clear()
        PriceTextBox.Clear()
        ExtendedPriceTextBox.Clear()
        With TitleTextBox
            .Clear()
            .Focus()
        End With
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        ' Exit the program.
        Me.Close()
    End Sub

    Private Sub SummaryToolStripMenuItem_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles SummaryToolStripMenuItem.Click
        ' Display the sales summary information.
        MessageBox.Show("Total sales: " & BookSale.AccumulatorSalesTotal.ToString("c"),
                        "Total Sales Amount",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation)
    End Sub
End Class
