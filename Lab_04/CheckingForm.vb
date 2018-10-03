'Project:     Lab 4
'Programmer:  Anthony DePinto
'Date:        Fall 2018
'Description: This project maintains a checking account balance.
'             The requested transaction is calculated and 
'             the new balance is displayed.
'             A summary includes all transactions.

Option Explicit On
Option Strict On

Public Class CheckingForm
    ' Declare variables
    Dim AccountBalanceDecimal As Decimal

    ' Accumulators
    Dim DepositAccumulatorInteger As Integer
    Dim DepositAmountAccumulatorDecimal As Decimal
    Dim CheckAccumulatorInteger As Integer
    Dim CheckAmountAccumulatordecimal As Decimal
    Dim ServiceChargeAccumulatorInteger As Integer
    Dim ServiceChargeAmountAccumulatorDecimal As Decimal

    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        'End the program

        Me.Close()
    End Sub


    Private Sub ClearTextBox_Click(sender As Object, e As EventArgs) Handles ClearTextBox.Click
        ' Zero out account balance and all accumulators
        AccountBalanceDecimal = 0
        DepositAccumulatorInteger = 0
        DepositAmountAccumulatorDecimal = 0
        CheckAccumulatorInteger = 0
        CheckAmountAccumulatordecimal = 0
        ServiceChargeAccumulatorInteger = 0
        ServiceChargeAmountAccumulatorDecimal = 0

        ' Update zeroed balance amount to balance label
        UpdateAccountBalanceField()
    End Sub

    Private Sub CalculateTextBox_Click(sender As Object, e As EventArgs) Handles CalculateTextBox.Click
        'Calculate the transaction and display the new balance.
        Dim AmountDecimal As Decimal

        If DepositRadioButton.Checked Or CheckRadioButton.Checked Or ChargeRadioButton.Checked Then
            Try
                AmountDecimal = Decimal.Parse(AmountTextBox.Text)

                If AmountDecimal < 0 Then
                    Throw New System.FormatException
                End If

                ' Calculate each transaction and keep track of summary information
                ' then display updated account balance to balance label
                AccountEvent(AmountDecimal)
                UpdateAccountBalanceField()
            Catch AmountException As FormatException
                MessageBox.Show("Please make sure that only positive numeric data has been entered.",
                    "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                With AmountTextBox
                    .Focus()
                    .SelectAll()
                End With
            Catch AnyException As Exception
                MessageBox.Show("Error: " & AnyException.Message)
            End Try
        Else
            MessageBox.Show("Please select deposit, check, or service charge", "Input needed")
        End If
    End Sub

    Sub AccountEvent(ByVal _AmountDecimal As Decimal)
        Select Case True
            Case DepositRadioButton.Checked
                DepositEvent(_AmountDecimal)
            Case CheckRadioButton.Checked
                CheckEvent(_AmountDecimal)
            Case ChargeRadioButton.Checked
                ChargeEvent()
            Case Else
                MessageBox.Show("Something weird happened",
                                "Weird error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Question)
        End Select
    End Sub

    Sub UpdateAccountBalanceField()
        ' Update label with current account balance decimal value
        BalanceTextBox.Text = AccountBalanceDecimal.ToString("c")
    End Sub

    Sub DepositEvent(ByVal __AmountDecimal As Decimal)
        ' Add __AmountDecimal to account balance
        AccountBalanceDecimal += __AmountDecimal

        ' Update running totals
        DepositAccumulatorInteger += 1
        DepositAmountAccumulatorDecimal += __AmountDecimal
    End Sub

    Sub CheckEvent(ByVal __AmountDecimal As Decimal)
        ' Check to make sure that the amount trying
        ' to be removed is not greater than account balance
        If __AmountDecimal <= AccountBalanceDecimal Then
            ' Amount to remove less than account balance so call withdraw subroutine
            ' to modify account balance variable
            Withdrawl(__AmountDecimal)

            ' Update accumulator values
            CheckAccumulatorInteger += 1
            CheckAmountAccumulatordecimal += __AmountDecimal
        Else
            ' Amount to remove greater than account balance so call service charge subroutine
            ' and dislay error message
            MessageBox.Show("Your balance is less than the amount you are trying to withdraw.",
                            "Overdraft Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
            ChargeEvent()
        End If
    End Sub

    Sub ChargeEvent()
        ' Declare constant service charge variable
        Const SERVICE_CHARGE_Decimal As Decimal = 10D

        ' Withdraw service charge from account balance
        Withdrawl(SERVICE_CHARGE_Decimal)

        ' Update running totals of service charge events and total service charges accrued
        ServiceChargeAccumulatorInteger += 1
        ServiceChargeAmountAccumulatorDecimal += SERVICE_CHARGE_Decimal
    End Sub

    Sub Withdrawl(ByVal __AmountDecimal As Decimal)
        AccountBalanceDecimal -= __AmountDecimal
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MessageBox.Show("Total number of deposits: " + DepositAccumulatorInteger.ToString + Environment.NewLine +
                        "Total amount of deposits: " + DepositAmountAccumulatorDecimal.ToString("c") + Environment.NewLine +
                        "Total number of checks: " + CheckAmountAccumulatordecimal.ToString + Environment.NewLine +
                        "Total Amount of checks: " + CheckAmountAccumulatordecimal.ToString("c") + Environment.NewLine +
                        "Total number of service charges: " + ServiceChargeAccumulatorInteger.ToString + Environment.NewLine +
                        "Total amount of service charges: " + ServiceChargeAmountAccumulatorDecimal.ToString("c"),
                        "Total account events",
                        MessageBoxButtons.OK)
    End Sub
End Class

