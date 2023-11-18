Public Class Form1
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim QZS, CS, AVE As Double
        QZS = (Val(TextBox1.Text) + Val(TextBox2.Text)) / 2
        CS = Val(TextBox3.Text)
        AVE = (QZS + CS + Val(TextBox4.Text)) / 3
        TextBox5.Text = AVE
        openCon()
        Try
            cmd.Connection = con
            cmd.CommandText = "INSERT INTO tbl_midterm (`quiz1`, `quiz2`, `classStanding`, `midtermExam`, `midtermGrade`) VALUES ('" & TextBox1.Text & "', 
            '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "')"
            cmd.ExecuteNonQuery()
            MsgBox("Succesfully added!")
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox10.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim QZS, CS, AVE As Double
        QZS = (Val(TextBox6.Text) + Val(TextBox7.Text)) / 2
        CS = Val(TextBox8.Text)
        AVE = (QZS + CS + Val(TextBox9.Text)) / 3
        TextBox10.Text = AVE
        openCon()
        Try
            cmd.Connection = con
            cmd.CommandText = "INSERT INTO tbl_finals (`quiz1`, `quiz2`, `classStanding`, `finalExam`, `finalGrade`) VALUES ('" & TextBox6.Text & "', 
            '" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & TextBox9.Text & "', '" & TextBox10.Text & "')"
            cmd.ExecuteNonQuery()
            MsgBox("Succesfully added!")
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Form2.Show()
        Form2.TextBox11.Text = TextBox5.Text
        Form2.TextBox12.Text = TextBox10.Text
        Dim OG As Double
        OG = (Val(TextBox5.Text) + Val(TextBox10.Text)) / 2
        Form2.TextBox13.Text = OG

        Select Case OG
            Case 0 To 74
                Form2.TextBox14.Text = "5.00"
            Case 75 To 76
                Form2.TextBox14.Text = "3.00"
            Case 77 To 79
                Form2.TextBox14.Text = "2.75"
            Case 80 To 82
                Form2.TextBox14.Text = "2.50"
            Case 83 To 85
                Form2.TextBox14.Text = "2.25"
            Case 86 To 88
                Form2.TextBox14.Text = "2.00"
            Case 89 To 91
                Form2.TextBox14.Text = "1.75"
            Case 92 To 94
                Form2.TextBox14.Text = "1.50"
            Case 95 To 97
                Form2.TextBox14.Text = "1.25"
            Case 98 To 100
                Form2.TextBox14.Text = "1.00"
        End Select

        If OG > 75 Then
            Form2.TextBox15.Text = "Passed!"
        Else
            Form2.TextBox15.Text = "Failed!"
        End If

        openCon()
        Try
            cmd.Connection = con
            cmd.CommandText = "INSERT INTO tbl_overall (`midtermGrade`, `finalGrade`, `overallGrade`, `equivGrade`, `remarks`) VALUES ('" & Form2.TextBox11.Text & "', 
            '" & Form2.TextBox12.Text & "', '" & Form2.TextBox13.Text & "', '" & Form2.TextBox14.Text & "', '" & Form2.TextBox15.Text & "')"
            cmd.ExecuteNonQuery()
            MsgBox("Succesfully added!")
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
