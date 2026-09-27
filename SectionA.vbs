Option Explicit

Class Account

'Class variables
Public strAccountNo
Public strAccountHolder
Public dblBalance

Private Sub Class_Initialize()

    dblBalance = 1000

End Sub

Public Function Withdrawal(dblAmount)

    dblBalance = dblBalance - dblAmount
    Withdrawal = dblBalance

End Function

Public Function Deposit(dblAmount)

    dblBalance = dblBalance + dblAmount
    Deposit = dblBalance

End Function

Public Function Transfer(strReceivingAccountNo, dblAmount)

    dblBalance = dblBalance - dblAmount
    Transfer = dblBalance

End Function

Public Function CheckBalance()

    MsgBox "Current Balance: R" & FormatNumber(dblBalance, 2)

    CheckBalance = dblBalance

End Function

Public Function getAccountNumber()

        getAccountNumber = "ACC-" & strAccountNo

End Function

Public Function getAccountDetails()

    getAccountDetails = "Account Holder: " & strAccountHolder & vbNewLine & _
                        vbNewLine & "Account Number: " & getAccountNumber() & vbNewLine & _
                        vbNewLine & "Remaining Balance: R" & _
                        FormatNumber(dblBalance, 2)

End Function

End Class

DIM objAccount
DIM strAccountNo
DIM strAccountHolder
DIM strOption
DIM strReceivingAccountNo
DIM strAmount
DIM dblAmount
DIM dblNewBalance
DIM strTitle

strTitle = "*******SmartBalance Banking System******"

Set objAccount = New Account

Do

   strAccountNo = Trim(InputBox( _
    "Please provide your account number for account verification.", _
    strTitle))

    
    If Len(strAccountNo) < 7 Or Len(strAccountNo) > 10 Or IsDigitsOnly(strAccountNo) = False Then

        MsgBox "Account number must contain 7 to 10 digits."

    End If

Loop While Len(strAccountNo) < 7 Or Len(strAccountNo) > 10 Or IsDigitsOnly(strAccountNo) = False

Do

    strAccountHolder = Trim(InputBox("Enter account holder name", strTitle))

    If strAccountHolder = "" Then

        MsgBox "Account holder name cannot be empty."

    End If

Loop While strAccountHolder = ""

objAccount.strAccountNo = strAccountNo
objAccount.strAccountHolder = strAccountHolder

Do

    strOption = Trim(InputBox( _
        "Select an option:" & vbNewLine & vbNewLine & _
        "1 - Withdrawal" & vbNewLine & _
        "2 - Deposit" & vbNewLine & _
        "3 - Transfer" & vbNewLine & _
        "4 - Check Balance" & vbNewLine & _
        "5 - Exit", _
        strTitle))

    Select Case strOption

        Case "1"

            Do

                strAmount = Trim(InputBox("Enter withdrawal amount", strTitle))

                If IsNumeric(strAmount) = True Then
                    dblAmount = CDbl(strAmount)
                Else
                    dblAmount = 0
                End If

                If dblAmount <= 0 Then

                    MsgBox "Withdrawal amount must be greater than zero."

                ElseIf dblAmount > objAccount.dblBalance Then

                    MsgBox "You cannot withdraw more than your current balance."

                End If

            Loop While dblAmount <= 0 Or dblAmount > objAccount.dblBalance

            dblNewBalance = objAccount.Withdrawal(dblAmount)

            MsgBox "Withdrawal successful." & vbNewLine & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "2"

            Do

                strAmount = Trim(InputBox("Enter deposit amount", strTitle))

                If IsNumeric(strAmount) = True Then
                    dblAmount = CDbl(strAmount)
                Else
                    dblAmount = 0
                End If

                If dblAmount <= 0 Then

                    MsgBox "Deposit amount must be greater than zero."

                End If

            Loop While dblAmount <= 0

            dblNewBalance = objAccount.Deposit(dblAmount)

            MsgBox "Deposit successful." & vbNewLine & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "3"

            Do

                strReceivingAccountNo = Trim(InputBox( _
                    "Enter receiving account number." & vbNewLine & _
                    "It must contain 7 to 10 digits.", _
                    strTitle))

                If Len(strReceivingAccountNo) < 7 Or Len(strReceivingAccountNo) > 10 Or IsDigitsOnly(strReceivingAccountNo) = False Then

                    MsgBox "Account number must contain 7 to 10 digits."

                End If

            Loop While Len(strReceivingAccountNo) < 7 Or Len(strReceivingAccountNo) > 10 Or IsDigitsOnly(strReceivingAccountNo) = False

            Do

                strAmount = Trim(InputBox("Enter transfer amount", strTitle))

                If IsNumeric(strAmount) = True Then
                    dblAmount = CDbl(strAmount)
                Else
                    dblAmount = 0
                End If

                If dblAmount <= 0 Then

                    MsgBox "Transfer amount must be greater than zero."

                ElseIf dblAmount > objAccount.dblBalance Then

                    MsgBox "You cannot transfer more than your current balance."

                End If

            Loop While dblAmount <= 0 Or dblAmount > objAccount.dblBalance

            dblNewBalance = objAccount.Transfer(strReceivingAccountNo, dblAmount)

            MsgBox "Transfer successful to ACC-" & strReceivingAccountNo & "." & _
                   vbNewLine & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "4"

            dblNewBalance = objAccount.CheckBalance()

            MsgBox objAccount.getAccountDetails()

        Case "5"

            MsgBox "Thank you. Program closed." & vbNewLine & vbNewLine & _
                   objAccount.getAccountDetails()

        Case Else

            MsgBox "Invalid option. Please select a number from 1 to 5."

    End Select

Loop While strOption <> "5"

Function IsDigitsOnly(strValue)

    DIM intPosition
    DIM strCharacter

    IsDigitsOnly = True

    For intPosition = 1 To Len(strValue)

        strCharacter = Mid(strValue, intPosition, 1)

        If InStr("0123456789", strCharacter) = 0 Then

            IsDigitsOnly = False
            Exit Function

        End If

    Next

End Function