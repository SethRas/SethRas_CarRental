'Seth Rasmussen
'RCET 0256
'https://github.com/SethRas/SethRas_CarRental.git

Option Explicit On
Option Strict Off
Option Compare Binary
Public Class SethRasRentalForm

    Dim ErrorMessage As String
    Dim DistanceMilesKilo As Boolean
    Dim Miles As Integer
    Dim DiscountSenior As String = ""
    Dim discountAAA As String = ""
    Dim cost As Integer
    Dim customerT, MilesT, ChargesT As Integer

    Private Sub SethRasRentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MileageChargeTextBox.Enabled = False
        DayChargeTextBox.Enabled = False
        TotalDiscountTextBox.Enabled = False
        TotalMilesTextBox.Enabled = False
        TotalChargeTextBox.Enabled = False
        MilesradioButton.Enabled = True
        Me.Text = "Car Rental"
    End Sub

    Function NameValidation(e As Boolean) As Boolean
        Select Case NameTextBox.Text
            Case ""
                e = True
                MsgBox("Enter Name")
                NameTextBox.Focus()
            Case Else
        End Select
        Return e
    End Function

    Function AddressValidation(e As Boolean) As Boolean
        Select Case AddressTextBox.Text
            Case ""
                e = True
                ErrorMessage += vbCrLf + "address Is not Valid"
                AddressTextBox.Focus()
            Case Else
                e = False
        End Select
        Return e
    End Function

    'City cannot be empty
    'City cannot be a number
    Function CityValidation(e As Boolean) As Boolean
        Select Case CityTextBox.Text
            Case ""
                e = True
                ErrorMessage += vbCrLf + "Enter a city"
                CityTextBox.Focus()
            Case Else
                Try
                    CityTextBox.Text = CInt(CityTextBox.Text)
                    e = True
                    ErrorMessage += vbCrLf + "city cannot include numbers"
                    CityTextBox.Focus()
                Catch ex As Exception
                    e = False
                End Try
                e = False
        End Select
        Return e
    End Function

    'State must be a name not a number
    Function StateValidation(e As Boolean) As Boolean
        Select Case StateTextBox.Text
            Case ""
                e = True
                ErrorMessage += vbCrLf + "Enter a state"
                StateTextBox.Focus()
            Case Else
                Try
                    StateTextBox.Text = CInt(StateTextBox.Text)
                    e = True
                    ErrorMessage += vbCrLf + "state cannot include numbers"
                    StateTextBox.Focus()
                Catch ex As Exception
                    e = False
                End Try
                e = False
        End Select
        Return e
    End Function

    'Zip Code has to be a number int
    Function zipCodeValidation(e As Boolean) As Boolean
        Try
            ZipCodeTextBox.Text = CInt(ZipCodeTextBox.Text)
            e = False
        Catch ex As Exception
            ZipCodeTextBox.Focus()
            e = True
            ErrorMessage += vbCrLf + "ZIP Code must be numerical"
        End Try
        Return e
    End Function

    'End of trip milage cannot be less than starting Feris
    Function OdometerValidation(e As Boolean)
        Try
            BeginOdometerTextBox.Text = CInt(BeginOdometerTextBox.Text)
            EndOdometerTextBox.Text = CInt(EndOdometerTextBox.Text)
            Miles = EndOdometerTextBox.Text - BeginOdometerTextBox.Text
            e = False
            If BeginOdometerTextBox.Text > EndOdometerTextBox.Text Then
                ErrorMessage += "Ending Milage Cannot Be less than the beginning milage"
                BeginOdometerTextBox.Focus()
                e = True
            End If
        Catch ex As Exception
            e = True
            ErrorMessage += vbCrLf + "Odometer input must be a number"
            BeginOdometerTextBox.Focus()
        End Try
        Return e
    End Function

    'No rentals for longer than 45 days
    'Days must be a number not days specific
    Function DayValidation(e As Boolean)
        Try
            DaysTextBox.Text = CInt(DaysTextBox.Text)
            e = False
            If DaysTextBox.Text > 45 Then
                ErrorMessage += "No rentals for longer than 45 days"
                e = True
                DaysTextBox.Focus()
            End If
        Catch ex As Exception
            e = True
            ErrorMessage += vbCrLf + "Days must be a number"
            DaysTextBox.Focus()
        End Try
        Return e
    End Function
    'miles under 200 are free
    'miles under 500 are 0.12
    'miles at more than 500 are at .10
    Function MilesValidation()

        Dim MilesA As Integer = 0
        Dim MilesB As Integer = 0
        Dim milesCost As Integer = 0

        If Miles > 200 Then
            MilesA = Miles - 200
            If MilesA > 500 Then
                MilesB = MilesA - 300
                milesCost += MilesB * 0.1
            End If
            If MilesA > 300 Then
                MilesA = 300
            End If
            milesCost += MilesA * 0.12

        Else
            milesCost = 0
        End If
        'convert to miles
        If DistanceMilesKilo = 0 Then
            milesCost = milesCost * 0.62
            Miles = 0.62 * Miles
        End If
        Return milesCost

    End Function

    'This function needs to check for discounts 
    '5% for AAA
    '3% for Senior
    Function DiscountCheck()
        Dim a, b, c As Integer
        If discountAAA = "AAA" Then
            a = cost * 0.05
        End If
        If DiscountSenior = "Senior" Then
            b += cost * 0.03
        End If
        c = a + b
        Return c
    End Function

    'Run the calculation, display in a box all information spaced out
    'Run the calculation of miles and discount
    Function CalculateCost()
        Dim e As Integer
        cost = 0
        cost += DaysTextBox.Text * 15
        cost += MilesValidation()
        cost = cost - DiscountCheck()
        e = MsgBox($"All information is as follows: " + vbCrLf +
                   $"Name: {NameTextBox.Text}" + vbCrLf +
                   $"Address: {AddressTextBox.Text}" + vbCrLf +
                   $"City: {CityTextBox.Text}" + vbCrLf +
                   $"State: {StateTextBox.Text}" + vbCrLf +
                   $"ZipCode: {ZipCodeTextBox.Text}" + vbCrLf +
                   $"{Miles} miles over {DaysTextBox.Text} day(s) using the {discountAAA + " " + DiscountSenior} discount applied" + vbCrLf +
                   $"With a total charge of ${cost}" + vbCrLf + vbCrLf +
                   "Please verify this information, if all looks to be correct,  Press yes to submit the rental agreement", MsgBoxStyle.Question.YesNo, "Submit Form")
        If e = 6 Then
            Summary()
        End If
        Return cost
    End Function

    'Clears inputs 
    Sub ClearAll()
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DayChargeTextBox.Text = ""
        DaysTextBox.Text = ""
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    'Uses clear all function
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        ClearAll()
    End Sub

    'Exit program
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub

    Sub displayTotal()
        MileageChargeTextBox.Text = Miles
        MileageChargeTextBox.Text = MilesValidation()
        DayChargeTextBox.Text = (DaysTextBox.Text * 15)
        TotalDiscountTextBox.Text = DiscountCheck()
        TotalChargeTextBox.Text = cost
    End Sub

    'Add all inputs into the summary and track customer
    Sub Summary()
        customerT += 1 + customerT
        MilesT += Miles + MilesT
        ChargesT += cost + ChargesT
    End Sub

    'Run all the validations if one is wrong throw an error
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim x As Boolean
        If DayValidation(x) Or OdometerValidation(x) Or zipCodeValidation(x) Or StateValidation(x) Or CityValidation(x) Or AddressValidation(x) Or NameValidation(x) = True Then
            MsgBox(ErrorMessage)

        Else
            CalculateCost()
            displayTotal()
        End If

        ErrorMessage = ""
    End Sub

    'Select Miles or Kilometers true or false in order to run the correct calculations
    Private Sub MilesRadiobutton_CheckedChanged(sender As Object, e As EventArgs) Handles MilesradioButton.CheckedChanged
        DistanceMilesKilo = 1
    End Sub
    Private Sub KilometersRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles KilometersradioButton.CheckedChanged
        DistanceMilesKilo = 0
    End Sub

    'Checks if discounts are selected
    Private Sub AAACheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles AAAcheckbox.CheckedChanged
        discountAAA = "AAA"
    End Sub
    Private Sub SeniorCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Seniorcheckbox.CheckedChanged
        DiscountSenior = "Senior"
    End Sub

    'Uses summary function put in a display box
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim er As Integer
        er = MsgBox($"Total Customers: {customerT}" + vbCrLf +
               $"Total Miles Driven: {MilesT} mi" + vbCrLf +
               $"Total Charges: ${ChargesT}" + vbCrLf +
               "Clear All?", MsgBoxStyle.Question.YesNo, "Clear")
        If er = 6 Then
            ClearAll()
            customerT = 0
            MilesT = 0
            ChargesT = 0
        End If
    End Sub
End Class
