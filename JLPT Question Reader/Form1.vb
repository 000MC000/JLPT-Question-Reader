Option Explicit On
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Runtime.InteropServices

Public Class Form1
    <DllImport("user32.dll")> Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Integer
    End Function

    Dim _progress As Integer
    Dim rAns As Integer = 100
    Dim Sr1 As StreamReader
    Dim SB As StringBuilder
    Dim cLine1 As String, cLine2 As String, cLine3 As String
    Dim questionNum As Integer
    Dim repeatQuestions As StringBuilder
    Delegate Sub _del1()





    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '#Load progress
        If Not File.Exists("JLPT Answers.txt") Then
            Dim fcstream As FileStream = New FileStream("JLPT Answers.txt", FileMode.Create)
            fcstream.Close()
            questionNum = 1
        Else
            Dim Ar1 As StreamReader = New StreamReader("JLPT Answers.txt")
            _progress = 0
            Dim _read As String = ""
            While Not _read Is Nothing
                _progress += 1
                _read = Ar1.ReadLine
            End While
            Ar1.Close()
            If _progress >= 1 Then
                questionNum = _progress
            Else
                questionNum = 1
            End If
        End If

        'Set repeatquestions to an empty string. 
        'When it is no longer an empty string, we can start to show only questions that need to be re-answered
        repeatQuestions = New StringBuilder("")

        'Read questions
        cLine1 = ""
        cLine2 = ""
        cLine3 = ""

        Dim SB As New StringBuilder("")


        Dim trd1 As New Threading.Thread(
            Sub()

                '##Read Text
                Sr1 = New StreamReader("JLPT1 Questions.txt")
                Dim findQuestion As Integer = 0
                While Not findQuestion = questionNum

                    cLine3 = cLine2
                    cLine2 = cLine1
                    cLine1 = Sr1.ReadLine

                    If cLine1.IndexOf("答案") <> -1 Then
                        findQuestion += 1
                    End If

                End While
                Sr1.Close()

                'Set the answer
                rAns = cLine1.Split(":"c)(1)


                'Start quiz (show text and enable buttons)
                Dim del2 As New _del1(AddressOf StartQuiz)
                Me.Invoke(del2)

            End Sub)
        trd1.Start()

        Form2.TopMost = True

    End Sub

    Private Sub StartQuiz()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False

        RichTextBox1.Text = "Question " & (cLine3.Split(" "c)(0) - 100) & ": " & Mid(cLine3, cLine3.IndexOf(cLine3.Split(" "c)(1)))
        setButtons()

    End Sub
    Private Sub setButtons()
        If cLine2.Contains("1)") = True Then
            Button1.Enabled = True

            Dim a As String
            a = Mid(cLine2, cLine2.IndexOf("1)") + 3)
            a = Mid(a, 1, a.IndexOf("2)"))
            Button1.Text = a
        End If

        If cLine2.Contains("2)") = True Then
            Button2.Enabled = True

            Dim a As String
            a = Mid(cLine2, cLine2.IndexOf("2)") + 3)
            If cLine2.Contains("3)") Then
                a = Mid(a, 1, a.IndexOf("3)"))
            End If
            Button2.Text = a
        End If

        If cLine2.Contains("3)") = True Then
            Button3.Enabled = True

            Dim a As String
            a = Mid(cLine2, cLine2.IndexOf("3)") + 3)
            If cLine2.Contains("4)") Then
                a = Mid(a, 1, a.IndexOf("4)"))
            End If
            Button3.Text = a
        End If

        If cLine2.Contains("4)") = True Then
            Button4.Enabled = True
            Button4.Text = Mid(cLine2, cLine2.LastIndexOf(")") + 2)
        End If
    End Sub
    Dim pressedbutton As Boolean = False
    Dim currentQ As Integer
    Dim repQBool As Boolean = False
    Private Sub LoadNextQuestion()
        'Reset the buttons
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False





        Dim _Sr1 As StreamReader = New StreamReader("JLPT1 Questions.txt")
        Dim findQuestion As Integer = 0

        ' Check if next question needs to be loaded from questions to repeat or not
        If repeatQuestions.ToString = "" Then

            '#############

            'Save the result of the previous question if necessary

            'If we are done repeating questions 
            If repQBool = False Then
                If returnedToNormal = False Then
                    SaveResult()
                Else
                    returnedToNormal = False
                End If

                While Not findQuestion = questionNum

                    cLine3 = cLine2
                    cLine2 = cLine1
                    cLine1 = _Sr1.ReadLine

                    If cLine1.IndexOf("答案") <> -1 Then
                        findQuestion += 1
                        rAns = cLine1.Split(":"c)(1)
                    End If

                End While

                _Sr1.Close()
            Else
                repQBool = False
                '################
                Dim _sr As StreamReader = New StreamReader("JLPT Answers.txt")
                Dim a As String = _sr.ReadLine
                Dim stb As StringBuilder = New StringBuilder("")

                While a IsNot Nothing
                    If a = currentQ & ":x" Then
                        stb.AppendLine(currentQ & ":" & answer)
                    Else
                        stb.AppendLine(a)
                    End If

                    a = _sr.ReadLine
                End While
                _sr.Close()
                My.Computer.FileSystem.WriteAllText("JLPT Answers.txt", stb.ToString, False)

                While Not findQuestion = questionNum

                    cLine3 = cLine2
                    cLine2 = cLine1
                    cLine1 = _Sr1.ReadLine

                    If cLine1.IndexOf("答案") <> -1 Then
                        findQuestion += 1
                        rAns = cLine1.Split(":"c)(1)
                    End If

                End While

                _Sr1.Close()
                '######################
            End If
        Else 'if we still have questions to repeat

            'replace the wrong answer with the new answer in answers.txt after a button press
            If pressedbutton = True Then
                Dim _sr As StreamReader = New StreamReader("JLPT Answers.txt")
                Dim a As String = _sr.ReadLine
                Dim stb As StringBuilder = New StringBuilder("")

                While a IsNot Nothing
                    If a = currentQ & ":x" Then
                        stb.AppendLine(currentQ & ":" & answer)
                    Else
                        stb.AppendLine(a)
                    End If

                    a = _sr.ReadLine
                End While
                _sr.Close()
                My.Computer.FileSystem.WriteAllText("JLPT Answers.txt", stb.ToString, False)
                pressedbutton = False
            ElseIf pressedbutton = False Then

            End If



            'move to the next question
            Dim int1 As Integer = repeatQuestions.ToString.Split(":")(0)
            While Not findQuestion = int1

                cLine3 = cLine2
                cLine2 = cLine1
                cLine1 = _Sr1.ReadLine

                If cLine1.IndexOf("答案") <> -1 Then
                    findQuestion += 1
                    rAns = cLine1.Split(":"c)(1)
                End If

            End While
            _Sr1.Close()

            currentQ = repeatQuestions.ToString.Split(":"c)(0)
            repeatQuestions = New StringBuilder(Mid(repeatQuestions.ToString, repeatQuestions.ToString.IndexOf(vbLf, 2) + 2))
            If repeatQuestions.ToString = "" Then
                repQBool = True
            End If
        End If



        RichTextBox1.Text = "Question " & (cLine3.Split(" "c)(0) - 100) & ": " & Mid(cLine3, cLine3.IndexOf(cLine3.Split(" "c)(1)))

        setButtons()
    End Sub
    Private Sub SaveResult()
        My.Computer.FileSystem.WriteAllText("JLPT Answers.txt", questionNum - 1 & ":" & answer & vbCrLf, True)
    End Sub
    Dim answer As String
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        pressedbutton = True
        If repQBool = False Then
            questionNum += 1
        End If

        If rAns = "4" Then
            answer = "o"
            Label1.ForeColor = Color.Green
            Label1.Text = "Correct"
        Else
            answer = "x"
            Label1.ForeColor = Color.Red
            Label1.Text = "Incorrect"
        End If

        '#load next question
        LoadNextQuestion()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        pressedbutton = True
        If repQBool = False Then
            questionNum += 1
        End If

        If rAns = "3" Then
            answer = "o"
            Label1.ForeColor = Color.Green
            Label1.Text = "Correct"
        Else
            answer = "x"
            Label1.ForeColor = Color.Red
            Label1.Text = "Incorrect"
        End If


        '#load next question
        LoadNextQuestion()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        pressedbutton = True
        If repQBool = False Then
            questionNum += 1
        End If

        If rAns = "2" Then
            answer = "o"
            Label1.ForeColor = Color.Green
            Label1.Text = "Correct"
        Else
            answer = "x"
            Label1.ForeColor = Color.Red
            Label1.Text = "Incorrect"
        End If


        '#load next question
        LoadNextQuestion()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pressedbutton = True
        If repQBool = False Then
            questionNum += 1
        End If

        If rAns = "1" Then
            answer = "o"
            Label1.ForeColor = Color.Green
            Label1.Text = "Correct"
        Else
            answer = "x"
            Label1.ForeColor = Color.Red
            Label1.Text = "Incorrect"
        End If


        '#load next question 

        LoadNextQuestion()
    End Sub


    Private Sub RichTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox1.KeyPress
        Select Case e.KeyChar.ToString
            Case "1", "a", "z"
                Button1_Click(Me, Nothing)
            Case "2", "s", "x"
                Button2_Click(Me, Nothing)
            Case "3", "d", "c"
                Button3_Click(Me, Nothing)
            Case "4", "f", "v"
                Button4_Click(Me, Nothing)
        End Select
    End Sub

    Private Sub RepeatMistakesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RepeatMistakesToolStripMenuItem.Click
        If My.Computer.FileSystem.FileExists("JLPT Answers.txt") Then

            repeatQuestions = New StringBuilder("")
            Dim Sr As StreamReader = New StreamReader("JLPT Answers.txt")
            Dim a As String = ""

            While a IsNot Nothing
                If a.IndexOf(":x") <> -1 Then
                    repeatQuestions.AppendLine(a)
                End If

                a = Sr.ReadLine
            End While
            Sr.Close()

            If repeatQuestions.ToString <> "" Then
                repQBool = False
                LoadNextQuestion()
            Else
                MsgBox("No questions to repeat.")
            End If

        Else
            MsgBox("No questions to repeat.")

        End If

    End Sub
    Dim returnedToNormal As Boolean = False
    Private Sub NormalModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalModeToolStripMenuItem.Click
        repeatQuestions.Clear()
        returnedToNormal = True
        LoadNextQuestion()
    End Sub
    Public clickedName2 As String
    Private Sub RichTextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles RichTextBox1.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If Form2.Visible = False AndAlso RichTextBox1.SelectedText.Length > 0 Then

                SetForegroundWindow(Form2.Handle)
                Form2.Left = MousePosition.X
                Form2.Top = MousePosition.Y
                Form2.Visible = True
                Form2.RichTextBox1.Text = "Searching translation..."

                Dim trd1 As New Threading.Thread(AddressOf doSearch)
                trd1.Start()

            ElseIf Form2.Visible = True Then
                Form2.Visible = False
            End If
        Else
            'If the left button is clicked
            Form2.Visible = False
        End If
    End Sub
    Private Sub doSearch()
        Try
            Dim EorJ As String
            clickedName2 = ""

            Dim del1 As New _del1(
                Sub()

                clickedName2 = RichTextBox1.SelectedText

            End Sub)
            Me.Invoke(del1)

            Dim sr As StreamReader = New StreamReader(WebRequest.Create(String.Format("http://honyakustar.com/search.php?searchTerm={0}", HttpUtility.UrlEncode(clickedName2.Replace("'s", " is").Replace(",", "").Replace(".", "").Replace("'t", " not").Replace(":", "").Replace(";", "").Replace("(", "").Replace(")", "").Replace("!", "").Replace("?", "")))).GetResponse.GetResponseStream)
            Dim string1 As String = sr.ReadToEnd
            Dim sAray As String() = string1.Split(vbLf)
            Dim translatedString As StringBuilder = New StringBuilder
            translatedString.Clear()
            Dim checkfirst As Boolean = True
            For i = 0 To sAray.Length - 1
                If sAray(i).IndexOf("english source") <> -1 AndAlso translatedString.Length < 400 Then
                    If checkfirst = True Then
                        EorJ = "e"
                        checkfirst = False
                    End If
                    translatedString.Append("---------" & vbCrLf)
                    Dim var As Integer = 0
                    translatedString.Append(Mid(sAray(i), 28, (sAray(i).Length - 4) - 28) & vbCrLf)
                    Do
                        var += 1
                        If sAray(i + var).IndexOf("</li>") <> -1 Then
                            translatedString.Append(Mid(sAray(i + var), 22, sAray(i + var).IndexOf("</li>") - 21) & vbCrLf)
                        End If
                        'Skip the lines we have already covered
                        i = i + var
                    Loop While sAray(i + var).IndexOf("</ul>") = -1
                ElseIf sAray(i).IndexOf("japanese source") <> -1 AndAlso translatedString.Length < 400 Then
                    If checkfirst = True Then
                        EorJ = "j"
                        checkfirst = False
                    End If
                    translatedString.Append("---------" & vbCrLf)
                    Dim var As Integer = 0
                    translatedString.Append(Mid(sAray(i), 29, (sAray(i).Length - 4) - 29) & vbCrLf)
                    Do
                        var += 1
                        If sAray(i + var).IndexOf("</li>") <> -1 Then
                            translatedString.Append(Mid(sAray(i + var), 21, sAray(i + var).IndexOf("</li>") - 20) & vbCrLf)
                        End If
                    Loop While sAray(i + var).IndexOf("</ul>") = -1
                    'Skip the lines we have already covered
                    i = i + var
                End If
            Next

            If translatedString.ToString.Length > 0 Then
                Dim del As New _del1(
                    Sub()

                    Form2.RichTextBox1.Text = clickedName2 & vbCrLf & Mid(translatedString.ToString, 1, translatedString.ToString.Length - 2)

                End Sub)
                Me.Invoke(del)
            Else
                Dim del As New _del1(
                    Sub()

                    Form2.RichTextBox1.Text = "No Results"

                End Sub)
                Me.Invoke(del)
            End If

        Catch ex As Exception
            Dim del As New _del1(
                Sub()
                    Form2.RichTextBox1.Text = "Failed to complete search: " & vbCrLf & "------" & vbCrLf & ex.ToString
                End Sub)
            Me.Invoke(del)
        End Try
    End Sub

    Private Sub RichTextBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles RichTextBox1.MouseMove
        If MousePosition.X - Form2.Left < -35 OrElse MousePosition.X - Form2.Left > 35 OrElse _
            MousePosition.Y - Form2.Top < -35 OrElse MousePosition.Y - Form2.Top > 35 Then

            Form2.Visible = False

        End If
    End Sub
End Class
