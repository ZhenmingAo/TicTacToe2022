' Fun TicTacToe project, updated as 2022 version
' Author: Zhenming Ao
Public Class TicTacToe2022

    Dim blnXO As Boolean = True 'Default starts off with "X"

    Private Sub btnBox1_Click(sender As Object, e As EventArgs) Handles btnBox1.Click
        If blnXO = True Then
            btnBox1.Text = "X"
            blnXO = False
        Else
            btnBox1.Text = "O"
            blnXO = True
        End If
        btnBox1.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox2_Click(sender As Object, e As EventArgs) Handles btnBox2.Click
        If blnXO = True Then
            btnBox2.Text = "X"
            blnXO = False
        Else
            btnBox2.Text = "O"
            blnXO = True
        End If
        btnBox2.Enabled = False
        checkForWinner()
    End Sub
    Private Sub btnBox3_Click(sender As Object, e As EventArgs) Handles btnBox3.Click
        If blnXO = True Then
            btnBox3.Text = "X"
            blnXO = False
        Else
            btnBox3.Text = "O"
            blnXO = True
        End If
        btnBox3.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox4_Click(sender As Object, e As EventArgs) Handles btnBox4.Click
        If blnXO = True Then
            btnBox4.Text = "X"
            blnXO = False
        Else
            btnBox4.Text = "O"
            blnXO = True
        End If
        btnBox4.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox5_Click(sender As Object, e As EventArgs) Handles btnBox5.Click
        If blnXO = True Then
            btnBox5.Text = "X"
            blnXO = False
        Else
            btnBox5.Text = "O"
            blnXO = True
        End If
        btnBox5.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox6_Click(sender As Object, e As EventArgs) Handles btnBox6.Click
        If blnXO = True Then
            btnBox6.Text = "X"
            blnXO = False
        Else
            btnBox6.Text = "O"
            blnXO = True
        End If
        btnBox6.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox7_Click(sender As Object, e As EventArgs) Handles btnBox7.Click
        If blnXO = True Then
            btnBox7.Text = "X"
            blnXO = False
        Else
            btnBox7.Text = "O"
            blnXO = True
        End If
        btnBox7.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox8_Click(sender As Object, e As EventArgs) Handles btnBox8.Click
        If blnXO = True Then
            btnBox8.Text = "X"
            blnXO = False
        Else
            btnBox8.Text = "O"
            blnXO = True
        End If
        btnBox8.Enabled = False
        checkForWinner()
    End Sub

    Private Sub btnBox9_Click(sender As Object, e As EventArgs) Handles btnBox9.Click
        If blnXO = True Then
            btnBox9.Text = "X"
            blnXO = False
        Else
            btnBox9.Text = "O"
            blnXO = True
        End If
        btnBox9.Enabled = False
        checkForWinner()
    End Sub

    Private Sub checkForWinner()
        'Horizontals checks
        If btnBox1.Text <> Nothing And btnBox1.Text = btnBox2.Text And btnBox2.Text = btnBox3.Text Then
            buttonDisable()
            btnBox1.BackColor = Color.Pink
            btnBox2.BackColor = Color.Pink
            btnBox3.BackColor = Color.Pink
            MessageBox.Show(btnBox1.Text + " is the winner!")
        ElseIf btnBox4.Text <> Nothing And btnBox4.Text = btnBox5.Text And btnBox5.Text = btnBox6.Text Then
            buttonDisable()
            btnBox4.BackColor = Color.Pink
            btnBox5.BackColor = Color.Pink
            btnBox6.BackColor = Color.Pink
            MessageBox.Show(btnBox4.Text + " is the winner!")
        ElseIf btnBox7.Text <> Nothing And btnBox7.Text = btnBox8.Text And btnBox8.Text = btnBox9.Text Then
            buttonDisable()
            btnBox7.BackColor = Color.Pink
            btnBox8.BackColor = Color.Pink
            btnBox9.BackColor = Color.Pink
            MessageBox.Show(btnBox7.Text + " is the winner!")
        End If

        'Verticals checks
        If btnBox1.Text <> Nothing And btnBox1.Text = btnBox4.Text And btnBox4.Text = btnBox7.Text Then
            buttonDisable()
            btnBox1.BackColor = Color.Pink
            btnBox4.BackColor = Color.Pink
            btnBox7.BackColor = Color.Pink
            MessageBox.Show(btnBox1.Text + " is the winner!")
        ElseIf btnBox2.Text <> Nothing And btnBox2.Text = btnBox5.Text And btnBox5.Text = btnBox8.Text Then
            buttonDisable()
            btnBox2.BackColor = Color.Pink
            btnBox5.BackColor = Color.Pink
            btnBox8.BackColor = Color.Pink
            MessageBox.Show(btnBox2.Text + " is the winner!")
        ElseIf btnBox3.Text <> Nothing And btnBox3.Text = btnBox6.Text And btnBox6.Text = btnBox9.Text Then
            buttonDisable()
            btnBox3.BackColor = Color.Pink
            btnBox6.BackColor = Color.Pink
            btnBox9.BackColor = Color.Pink
            MessageBox.Show(btnBox3.Text + " is the winner!")
        End If

        'Diagonal checks
        If btnBox1.Text <> Nothing And btnBox1.Text = btnBox5.Text And btnBox5.Text = btnBox9.Text Then
            buttonDisable()
            btnBox1.BackColor = Color.Pink
            btnBox5.BackColor = Color.Pink
            btnBox9.BackColor = Color.Pink
            MessageBox.Show(btnBox1.Text + " is the winner!")
        ElseIf btnBox3.Text <> Nothing And btnBox3.Text = btnBox5.Text And btnBox5.Text = btnBox7.Text Then
            buttonDisable()
            btnBox3.BackColor = Color.Pink
            btnBox5.BackColor = Color.Pink
            btnBox7.BackColor = Color.Pink
            MessageBox.Show(btnBox3.Text + " is the winner!")
        End If
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        blnXO = True
        btnBox1.Text = Nothing
        btnBox2.Text = Nothing
        btnBox3.Text = Nothing
        btnBox4.Text = Nothing
        btnBox5.Text = Nothing
        btnBox6.Text = Nothing
        btnBox7.Text = Nothing
        btnBox8.Text = Nothing
        btnBox9.Text = Nothing
        btnBox1.Enabled = True
        btnBox2.Enabled = True
        btnBox3.Enabled = True
        btnBox4.Enabled = True
        btnBox5.Enabled = True
        btnBox6.Enabled = True
        btnBox7.Enabled = True
        btnBox8.Enabled = True
        btnBox9.Enabled = True
        btnBox1.BackColor = DefaultBackColor
        btnBox2.BackColor = DefaultBackColor
        btnBox3.BackColor = DefaultBackColor
        btnBox4.BackColor = DefaultBackColor
        btnBox5.BackColor = DefaultBackColor
        btnBox6.BackColor = DefaultBackColor
        btnBox7.BackColor = DefaultBackColor
        btnBox8.BackColor = DefaultBackColor
        btnBox9.BackColor = DefaultBackColor
    End Sub
    Private Sub buttonDisable()
        btnBox1.Enabled = False
        btnBox2.Enabled = False
        btnBox3.Enabled = False
        btnBox4.Enabled = False
        btnBox5.Enabled = False
        btnBox6.Enabled = False
        btnBox7.Enabled = False
        btnBox8.Enabled = False
        btnBox9.Enabled = False
    End Sub

End Class
