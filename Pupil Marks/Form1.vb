'Andrew Luca Simpson
'1/9/23
'Pupile Marks

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim pupilName(20) As String
        Dim prelimMark(20) As Integer
        Dim courseMark(20) As Integer
        Dim percentage(20) As Single
        Dim topPosition As Single

        getResults(pupilName, prelimMark, courseMark)
        calculatePercentages(prelimMark, courseMark, percentage, topPosition)
        FindTopMark(percentage, topPosition)
        displayResults(pupilName, topPosition, percentage)

    End Sub

    Sub getResults(ByRef pupilname() As String, ByRef prelimMark() As Integer, ByRef courseMark() As Integer)

        Dim filename As String

        filename = "H:\Visual Studio\Higher\pupilMarks.csv"

        FileOpen(1, filename, OpenMode.Input)

        For i = 0 To 19

            Input(1, pupilname(i))
            Input(1, prelimMark(i))
            Input(1, courseMark(i))

        Next

        FileClose(1)

    End Sub

    Sub calculatePercentages(ByVal prelimMark() As Integer, ByVal courseMark() As Integer, ByRef percentage() As Single, ByRef topPosition As String)

        For i = 0 To 19

            percentage(i) = prelimMark(i) + courseMark(i) / 1.5
            percentage(i) = Math.Round(percentage(i), 2)

        Next

        topPosition = FindTopMark(percentage, topPosition)

    End Sub

    Function FindTopMark(ByVal percentage() As Single, ByRef topPosition As Single)

        topPosition = percentage(0)
        For i = 0 To 19

            If percentage(i) > topPosition Then
                topPosition = percentage(i)
            End If
        Next

        Return topPosition

    End Function

    Sub displayResults(ByVal PupileName() As String, ByVal topPosition As Single, ByVal percentage() As Single)

        For i = 0 To 19
            If topPosition = percentage(i) Then
                ListBox1.Items.Add("Top Pupil is " & PupileName(topPosition) & "with " & percentage(topPosition))
            End If
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
