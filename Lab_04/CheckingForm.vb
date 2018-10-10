'Project:     Lab 4
'Programmer:  Anthony DePinto
'Date:        Fall 2018
'Description: This project maintains a checking account balance.
'             The requested transaction is calculated and 
'             the new balance is displayed.
'             A summary includes all transactions.


' Date: 3 October 2018
' Author: Keith Smith

Option Explicit On
Option Strict On

Public Class CheckingForm
    ' Declare variables
    Dim AccountBalanceDecimal As Decimal

    ' Accumulators for account events and amounts
    Dim DepositAccumulatorInteger As Integer
    Dim DepositAmountAccumulatorDecimal As Decimal
    Dim CheckAccumulatorInteger As Integer
    Dim CheckAmountAccumulatordecimal As Decimal
    Dim ServiceChargeAccumulatorInteger As Integer
    Dim ServiceChargeAmountAccumulatorDecimal As Decimal

    ' Closes program
    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        'End the program
        Me.Close()
    End Sub

    ' Zeroes out all accumulator values and resets account balance and label text
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
        BalanceTextBox.Text = AccountBalanceDecimal.ToString("c")

        ' Clear amount text box entry field
        AmountTextBox.Clear()
    End Sub

    'Calculate the transaction and display the new balance.
    Private Sub CalculateTextBox_Click(sender As Object, e As EventArgs) Handles CalculateTextBox.Click
        ' Declare temporary decimal to hold amount for account event
        Dim AmountDecimal As Decimal

        If DepositRadioButton.Checked Or CheckRadioButton.Checked Or ChargeRadioButton.Checked Then
            Try
                AmountDecimal = Decimal.Parse(AmountTextBox.Text)
                ' Make sure amount is positive (no negative values)
                If AmountDecimal < 0 Then
                    Throw New System.FormatException
                End If

                ' Calculate each transaction and keep track of summary information
                AccountEvent(AmountDecimal)
                UpdateAccountBalanceLabel()
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

    ' Update account balance label text property
    Sub UpdateAccountBalanceLabel()
        BalanceTextBox.Text = AccountBalanceDecimal.ToString("c")
    End Sub

    ' Account event subroutine, takes in an amount and checks
    ' to see which radio button has been checked.
    Sub AccountEvent(ByVal _AmountDecimal As Decimal)
        ' Figure out which radio button was selected and perform appropriate action
        Select Case True
            Case DepositRadioButton.Checked
                DepositEvent(_AmountDecimal)
            Case CheckRadioButton.Checked
                CheckEvent(_AmountDecimal)
            Case ChargeRadioButton.Checked
                ChargeEvent(_AmountDecimal)
            Case Else
                MessageBox.Show("Something weird happened",
                                "Weird error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Question)
        End Select

        ' After event, update account balance label text property
    End Sub

    ' Basic deposit event, add amount to account balance
    Sub DepositEvent(ByVal __AmountDecimal As Decimal)
        ' Add __AmountDecimal to account balance
        AccountBalanceDecimal += __AmountDecimal

        ' Update running totals
        DepositAccumulatorInteger += 1
        DepositAmountAccumulatorDecimal += __AmountDecimal
    End Sub

    ' Basic withdraw event, remove amount from balance
    Sub WithdrawEvent(ByVal __AmountDecimal As Decimal)
        ' Subtract amount decimal from account balance
        AccountBalanceDecimal -= __AmountDecimal

        ' Updating accumulators happens one level up since
        ' this routine is used by the check and the service charge subroutines
    End Sub

    ' check event subroutine
    Sub CheckEvent(ByVal __AmountDecimal As Decimal)

        ' Check to make sure that the amount trying
        ' to be removed is not greater than account balance
        If __AmountDecimal <= AccountBalanceDecimal Then
            ' Amount to remove less than account balance so call withdraw subroutine
            ' to modify account balance variable
            WithdrawEvent(__AmountDecimal)

            ' Update accumulator values
            CheckAccumulatorInteger += 1
            CheckAmountAccumulatordecimal += __AmountDecimal
        Else
            ' Amount to remove greater than account balance so call service charge subroutine
            ' and dislay error message
            MessageBox.Show("Insufficient funds: $10 service charge.",
                            "Overdraft Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)

            ChargeEvent()
        End If
    End Sub

    ' Event for service charge, withdraws amount and updates
    ' accumulator values.
    Sub ChargeEvent(Optional __AmountDecimal As Decimal = 10D)
        ' Remove amount from account balance regardless if it makes
        ' the account balance negative
        WithdrawEvent(__AmountDecimal)

        ' Update accumulator values
        ServiceChargeAccumulatorInteger += 1
        ServiceChargeAmountAccumulatorDecimal += __AmountDecimal
    End Sub

    ' Subroutine to display running totals of all account events and amounts
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

    Private Sub CheckingForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdateAccountBalanceLabel()
    End Sub
End Class

