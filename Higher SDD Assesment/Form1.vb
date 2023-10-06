'Andrew Luca Simpson
'29/09/2023
'Higher SDD Assesment Step 7
'SCN: 211226501

Public Class Form1

    Structure details
        Dim PupilName As String 'Declares PupilName as a character and stores in a unique location memory
        Dim CourseWorkMark As Integer 'Declares CourseWorkMark as a whole number and stores in a unique location memory
        Dim PrelimMark As Integer 'Declares PrelimMark as a whole number and stores in a unique location memory
    End Structure

    Dim classmarks(15) As details

    Dim PupilName(15) As String 'Declares PupilName as a character and stores in a unique location memory
    Dim CourseWorkMark(15) As Integer 'Declares CourseWorkMark as a whole number and stores in a unique location memory
    Dim PrelimMark(15) As Integer 'Declares PrelimMark as a whole number and stores in a unique location memory
    Dim TotalMark(15) As Integer 'Declares TotalMark as a whole number and stores in a unique location memory
    Dim Percentage(15) As Integer 'Declares Precentage as a whole number and stores in a unique location memory
    Dim Grade(15) As String 'Declares Grade as a character and stores in a unique location memory
    Dim AOcurence As Integer 'Declares Precentage as a whole number and stores in a unique location memory

    Private Sub BtnStart_Click(sender As Object, e As EventArgs) Handles BtnStart.Click

        GetDetails(classmarks, PupilName, CourseWorkMark, PrelimMark) 'This subroutine get the details
        CalculatePercentage(TotalMark, Percentage, CourseWorkMark, PrelimMark) 'This subroutine calculates the percentage
        CalculateGrade(Percentage, Grade) 'This subroutine calculates the grade
        countOcurrence() 'This subroutine finds how many A grades there are
        FindMax(Percentage, classmarks, PupilName)
        DisplayResults(PupilName, Percentage, Grade) 'This subroutine displays the resulats to the user in a listbox

    End Sub 'Ends the subroutine

    Sub GetDetails(ByRef classmarks() As details, ByRef PupilName() As String, ByRef CourseWorkMark() As Integer, ByRef PrelimMark() As Integer) 'This subroutine get the details from the user

        Dim filename As String

        filename = "H:\Visual Studio\SDD Assessment\classmarks.csv"

        FileOpen(1, filename, OpenMode.Input)

        For i = 0 To 14

            Input(1, classmarks(i).PupilName)
            Input(1, classmarks(i).CourseWorkMark)
            Input(1, classmarks(i).PrelimMark)

        Next

        FileClose(1)

    End Sub 'Ends the subroutine

    Sub CalculatePercentage(ByRef TotalMark() As Integer, ByRef Percentage() As Integer, ByVal CourseWorkMark() As Integer, ByVal PrelimMark() As Integer) 'This subroutine calculates the percentage

        For i = 0 To 14

            TotalMark(i) = classmarks(i).CourseWorkMark + classmarks(i).PrelimMark 'This adds the CourseWorkMark and the PrelimMark together to get TotalMark

            Percentage(i) = TotalMark(i) * 100 \ 150 'This takes the TotalMark and times it by 100 then divides it by 150 to get the percentage

        Next

    End Sub 'Ends the subroutine

    Sub CalculateGrade(ByVal Percentage() As Integer, ByRef Grade() As String) 'This subroutine calculates the grade

        For i = 0 To 14

            If Percentage(i) >= 70 Then 'Starts an if statment
                Grade(i) = "A"
            ElseIf Percentage(i) >= 60 And Percentage(i) <= 69 Then
                Grade(i) = "B"
            ElseIf Percentage(i) >= 50 And Percentage(i) <= 59 Then
                Grade(i) = "C"
            ElseIf Percentage(i) >= 45 And Percentage(i) <= 49 Then
                Grade(i) = "D"
            Else
                Grade(i) = "No Grade"

            End If 'Ends the if statment

        Next

    End Sub 'Ends the subroutine

    Sub countOcurrence()

        For i = 0 To 14

            If Percentage(i) >= 70 Then 'Starts an if statment
                AOcurence = AOcurence + 1 'Adds 1 to AOcurence
            End If 'Ends the if statment

        Next

        LstDisplay.Items.Add("There are " & AOcurence & " A Grades") 'Sends text and AOcurence to display
        LstDisplay.Items.Add(" ")

    End Sub

    Sub FindMax(ByVal Percentage() As Integer, ByVal classmarks() As details, ByVal PupilName() As String)

        Dim max As Integer 'Declares max as a whole number and stores in a unique location memory
        Dim position As Integer 'Declares position as a whole number and stores in a unique location memory

        max = Percentage(0)

        For i = 1 To 14

            If Percentage(i) > max Then
                max = Percentage(i)
                position = i
            End If

        Next

        LstDisplay.Items.Add("The pupil with the top mark is " & classmarks(position).PupilName & " with " & Percentage(position) & "%") 'Sends text, PupilName and percentage to display
        LstDisplay.Items.Add(" ")

    End Sub
    Sub DisplayResults(ByVal PupilName() As String, ByVal Percentage() As Integer, ByVal Grade() As String) 'This subroutine displays the resulats to the user in a listbox

        For i = 0 To 14

            LstDisplay.Items.Add(classmarks(i).PupilName) 'Sends PupilName to display
            LstDisplay.Items.Add("Percentage is " & Percentage(i) & "%") 'Sends text and percentage to display
            LstDisplay.Items.Add("Grade is " & Grade(i)) 'Sends text and Grade to display
            LstDisplay.Items.Add(" ")

        Next

    End Sub 'Ends the subroutine

    Private Sub BtnEnd_Click(sender As Object, e As EventArgs) Handles BtnEnd.Click
        Me.Close()
    End Sub 'Ends the subroutine
End Class
