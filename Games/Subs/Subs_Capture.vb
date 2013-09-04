Module Subs_Capture
    Public Sub Capture_this()
        Dim new_pos As Point
        Do
            new_pos.X = Rnd() * Form1.Width
            new_pos.Y = Rnd() * Form1.Height
        Loop While (new_pos.X > Form1.Width - 55) Or (new_pos.Y > Form1.Height - 100) Or (new_pos.X < 110 And new_pos.Y < 110)

        Form1.Cap.Location = new_pos
    End Sub


    Public Sub difficulty_select()
        If Form1.RadioButton6.Checked Then
            Form1.Timer1.Interval = 300
            Form1.Cap.Text = "AAA"
        ElseIf Form1.RadioButton5.Checked Then
            Form1.Timer1.Interval = 150
            Form1.Cap.Text = "AA"
        ElseIf Form1.RadioButton4.Checked Then
            Form1.Timer1.Interval = 75
            Form1.Cap.Text = "A"
        End If
    End Sub

    Public Sub Capture_Initialize(ByRef is_captured As Boolean)
        is_captured = False
        Form1.Timer1.Start()
        Form1.Timer1.Interval = 300
        Form1.Cap.Text = "AAA"
        difficulty_select()
    End Sub
End Module
