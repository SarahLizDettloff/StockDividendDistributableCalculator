Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GroupBox2.Hide()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtAnnualDividendperShare.Text = ""
        txtDividendYieldCalc.Text = ""
        txtNumberofShares.Text = ""
        txtParValue.Text = ""
        txtPriceperShare.Text = ""
        txtResultStockDividendDistri.Text = ""
        txtStockDividendPercentage.Text = ""
        cbSharesNumber.Text = ""
        GroupBox2.Hide()
        checkboxHelp.Checked = False

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        MessageBox.Show("
A stock dividend is important because it is a dividend payment made in the form of additional shares rather than an actual cash payout. Many companies who are short on liquid cash will distribute this dividend to shareholders. In other words, stocks are distributed among the shareholders. Many investors find this appealing because it presents the oppertunity to generate equity.
" + vbCrLf + "This can be calculated by using the following equation: " + vbCrLf + "
Stock Dividend Distributable = (Stock Dividend Percentage / 100) x (Number of shares) x (Par Value)" + vbCrLf + "
Stock dividend percentage is the percentage one will earn off of the stock. " + vbCrLf + "
This can be calculated by dividing the annual dividend share by the price of the shares.", "About Stock Dividends")
    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Dim NumberShares As Double = 0
        Dim Value As Double = 0
        Dim dividYield As Double = 0
        Dim ShareType As Double = 0
        Dim Share As Double = 0
        Dim StockDividendDistributable As Double = 0

        If Double.TryParse(txtNumberofShares.Text, NumberShares) = 0 Then
            MessageBox.Show("Please enter the number of shares.", "Input Error")
        End If
        If Double.TryParse(txtParValue.Text, Value) = 0 Then
            MessageBox.Show("Please enter the par value.", "Input Error")
        End If
        If Double.TryParse(txtStockDividendPercentage.Text, dividYield) = 0 Then
            MessageBox.Show("Please enter the percentage of stock dividend..", "Input Error")
        End If

        If Not IsNumeric(NumberShares) Then
            MessageBox.Show("Please enter the number of shares.", "Input Error")
            txtNumberofShares.Focus()
        End If

        If cbSharesNumber.SelectedIndex = 0 Then 'Hundred
                ShareType = 100
            ElseIf cbSharesNumber.SelectedIndex = 1 Then 'Thousand
                ShareType = 1000
            ElseIf cbSharesNumber.SelectedIndex = 2 Then 'Ten Thousand
                ShareType = 10000
            ElseIf cbSharesNumber.SelectedIndex = 3 Then 'Hundred Thousand
                ShareType = 100000
            ElseIf cbSharesNumber.SelectedIndex = 4 Then 'Million
                ShareType = 1000000
            ElseIf cbSharesNumber.SelectedIndex = 5 Then 'Ten Million
                ShareType = 10000000
            ElseIf cbSharesNumber.SelectedIndex = 6 Then 'Hundred Million
                ShareType = 100000000
            ElseIf cbSharesNumber.SelectedIndex = 7 Then 'Trillion
                ShareType = 1000000000
            Else
                MessageBox.Show("Please select a share type.", "Input error")
            cbSharesNumber.Focus()
        End If

        Share = ShareType * NumberShares

        StockDividendDistributable = ((dividYield / 100) * Share * Value)
        txtResultStockDividendDistri.Text = FormatCurrency(StockDividendDistributable)
    End Sub

    Private Sub btnCalcDividendYield_Click(sender As Object, e As EventArgs) Handles btnCalcDividendYield.Click
        Dim dividYield As Double = 0
        Dim AnnualDividendperShare As Double = 0
        Dim PriceperShare As Double = 0

        If Double.TryParse(txtPriceperShare.Text, PriceperShare) = 0 Then
            MessageBox.Show("Please enter the annual dividend per a share.", "Input Error")
            txtPriceperShare.Focus()
        End If
        If Double.TryParse(txtAnnualDividendperShare.Text, AnnualDividendperShare) = 0 Then
            MessageBox.Show("Please enter the price per a share..", "Input Error")
            txtAnnualDividendperShare.Focus()
        End If

        dividYield = ((AnnualDividendperShare / PriceperShare) * 100)
        txtDividendYieldCalc.Text = dividYield
        txtStockDividendPercentage.Text = dividYield
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub checkboxHelp_CheckedChanged(sender As Object, e As EventArgs) Handles checkboxHelp.CheckedChanged
        If checkboxHelp.Checked = True Then
            GroupBox2.Show()
        End If
        If checkboxHelp.Checked = False Then
            GroupBox2.Hide()
        End If
    End Sub

End Class
