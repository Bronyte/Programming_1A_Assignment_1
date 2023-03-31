Public Class Form1

    'Declaring variables.

    '>>Inputs
    Public mark_Prg2D As String
    Public mark_ArtInt As String
    Public mark_DgtLgc As String
    Public mark_PrgC As String
    Public mark_PrgJva As String
    Public mark_DtBs As String

    '>>Outputs
    Public mark_total As Integer
    Public average As Integer
    Public grade As String
    Public decision As String

    '>>Other
    Public application_is_active As Boolean = True
    Public marks_string_array(5) As String
    Public marks_integer_array(5) As Integer
    Public boolean_grade_array(5) As Integer
    Public Decision_val As Integer

    'Display output when clicked
    Private Sub Button_calculate_Click(sender As Object, e As EventArgs) Handles Button_calculate.Click
        ClearOutput()

        If (isCorrectFormat(TextBox_Prg2D.Text)) Then
            mark_Prg2D = TextBox_Prg2D.Text
        End If

        If (isCorrectFormat(TextBox_ArtInt.Text)) Then
        End If

        mark_ArtInt = TextBox_ArtInt.Text
        If (isCorrectFormat(TextBox_DgtLgc.Text)) Then
            mark_DgtLgc = TextBox_DgtLgc.Text
        End If

        If (isCorrectFormat(TextBox_PrgC.Text)) Then
            mark_PrgC = TextBox_PrgC.Text
        End If

        If (isCorrectFormat(TextBox_PrgJva.Text)) Then
            mark_PrgJva = TextBox_PrgJva.Text
        End If

        If (isCorrectFormat(TextBox_DtBs.Text)) Then
            mark_DtBs = TextBox_DtBs.Text
        End If


        marks_integer_array = {mark_Prg2D, mark_ArtInt, mark_DgtLgc, mark_PrgC, mark_PrgJva, mark_DtBs}
        'Print Total.
        TextBox_total.Text = myTotal(marks_integer_array)

        'Calculating average.       xbar = sum of x / n
        average = myTotal(marks_integer_array) / 6

        'Print Average.
        TextBox_average.Text = average

        'Print Grade.
        TextBox_grade.Text = myGrade(average)

        'Print Decision.
        TextBox_decision.Text = isPass(marks_integer_array)

    End Sub

    'Close forms when clicked.
    Private Sub Button_exit_Click(sender As Object, e As EventArgs) Handles Button_exit.Click
        ApplicationExit()
    End Sub

    Public Sub ApplicationExit()
        Application.Exit()
    End Sub

    'Clear all entries when clicked
    Private Sub Button_enter_new_details_Click(sender As Object, e As EventArgs) Handles Button_enter_new_details.Click
        ClearInput()
        ClearOutput()
    End Sub

    Public Sub ClearInput()
        'Clearing all TextBox inputs.
        TextBox_name.Clear()

        TextBox_Prg2D.Clear()
        TextBox_ArtInt.Clear()
        TextBox_DgtLgc.Clear()
        TextBox_PrgC.Clear()
        TextBox_PrgJva.Clear()
        TextBox_DtBs.Clear()
    End Sub

    Public Sub ClearOutput()
        'Clearing all TextBox outputs.
        TextBox_total.Clear()
        TextBox_average.Clear()
        TextBox_grade.Clear()
        TextBox_decision.Clear()

        'Clearing all arrays.
        ReDim marks_integer_array(5)
        ReDim boolean_grade_array(5)
    End Sub

    'Exception handling. Check if input is in the correct format (integer), and within range (0-100).
    Private Function isCorrectFormat(mark As String) As Boolean
        Dim temp As Integer

        If (String.IsNullOrEmpty(mark)) Then
            MessageBox.Show("Empty field: Please confirm your input")
            Return False
        ElseIf (Not Integer.TryParse(mark, temp)) Then
            MessageBox.Show("Incorrect Format: Please confirm your input")
            Return False
        ElseIf mark < 0 Or mark > 100 Then
            MessageBox.Show(("Input out of range. Please check your input."))
            Return False
        Else
            Return True
        End If
    End Function

    Private Function myTotal(marks_integer_array As Integer()) As Integer
        mark_total = 0

        For i As Integer = 1 To 6
            'Summing all values in the marks_integer_array().
            mark_total += marks_integer_array(i - 1)
        Next

        Return mark_total
    End Function

    Private Function myGrade(average As Integer) As String
        'Determining Grade by comparing the Average.
        Select Case average
            Case 0 To 49
                grade = "Fail"
            Case 50 To 59
                grade = "Pass"
            Case 60 To 69
                grade = "Credit"
            Case 70 To 74
                grade = "Merit"
            Case 75 To 100
                grade = "Distinction"
            Case Else
                grade = "Mark out of range"
        End Select
        Return grade
    End Function

    Public Function isPass(marks_integer_array As Integer()) As String
        Decision_val = 0
        For i As Integer = 1 To 6

            'Pass = 0, Fail = 1
            If marks_integer_array(i - 1) >= 0 And marks_integer_array(i - 1) <= 49 Then
                Decision_val += 1
            End If

            'Determine Decision based on the total value of Decision_val.
            Select Case Decision_val
                Case 0
                    decision = "Proceed"
                Case 1
                    decision = "Proceed, carrying 1 Failed subject"
                Case 2
                    decision = "Proceed, carrying 2 Failed subject"
                Case 3
                    decision = "Proceed, carrying 3 Failed subject"
                Case 4
                    decision = "Repeat Semester"
                Case Else
                    decision = "N/A"
            End Select
        Next
        Return decision
    End Function
End Class
