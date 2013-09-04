Module Subs_Sudoku
    Public Sub Sudoku_Initialize(ByRef sudoku_box(,) As Byte)
        Form1.Timer1.Stop()
        clear_box(sudoku_box)
        read_in(sudoku_box)

        Form1.Timer1.Interval = 10
        Form1.Timer1.Start()

    End Sub

    Public Sub Sudoku(ByRef sudoku_box(,) As Byte)
        read_in(sudoku_box)
    End Sub

    'Checks and puts input from textboxes into sudoku_box
    Private Sub put_in_box(ByRef sudoku_box As Byte, ByRef Text As Object)
        If (Text.Text = vbNullChar) Then
            sudoku_box = 0
        Else
            sudoku_box = Microsoft.VisualBasic.Val(Text.Text)
            If (sudoku_box < 1) Or (sudoku_box > 9) Then
                sudoku_box = 0
                Text.clear()
            End If
        End If
    End Sub
    'This sub clears the array sudoku_box
    Private Sub clear_box(ByRef sudoku_box(,) As Byte)
        '1 to 9
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.TextBox3.Clear()
        Form1.TextBox4.Clear()
        Form1.TextBox5.Clear()
        Form1.TextBox6.Clear()
        Form1.TextBox7.Clear()
        Form1.TextBox8.Clear()
        Form1.TextBox9.Clear()
        '10 to 19
        Form1.TextBox10.Clear()
        Form1.TextBox11.Clear()
        Form1.TextBox12.Clear()
        Form1.TextBox13.Clear()
        Form1.TextBox14.Clear()
        Form1.TextBox15.Clear()
        Form1.TextBox16.Clear()
        Form1.TextBox17.Clear()
        Form1.TextBox18.Clear()
        Form1.TextBox19.Clear()
        '20 to 29
        Form1.TextBox20.Clear()
        Form1.TextBox21.Clear()
        Form1.TextBox22.Clear()
        Form1.TextBox23.Clear()
        Form1.TextBox24.Clear()
        Form1.TextBox25.Clear()
        Form1.TextBox26.Clear()
        Form1.TextBox27.Clear()
        Form1.TextBox28.Clear()
        Form1.TextBox29.Clear()
        '30 to 39
        Form1.TextBox30.Clear()
        Form1.TextBox31.Clear()
        Form1.TextBox32.Clear()
        Form1.TextBox33.Clear()
        Form1.TextBox34.Clear()
        Form1.TextBox35.Clear()
        Form1.TextBox36.Clear()
        Form1.TextBox37.Clear()
        Form1.TextBox38.Clear()
        Form1.TextBox39.Clear()
        '40 to 49
        Form1.TextBox40.Clear()
        Form1.TextBox41.Clear()
        Form1.TextBox42.Clear()
        Form1.TextBox43.Clear()
        Form1.TextBox44.Clear()
        Form1.TextBox45.Clear()
        Form1.TextBox46.Clear()
        Form1.TextBox47.Clear()
        Form1.TextBox48.Clear()
        Form1.TextBox49.Clear()
        '50 to 59
        Form1.TextBox50.Clear()
        Form1.TextBox51.Clear()
        Form1.TextBox52.Clear()
        Form1.TextBox53.Clear()
        Form1.TextBox54.Clear()
        Form1.TextBox55.Clear()
        Form1.TextBox56.Clear()
        Form1.TextBox57.Clear()
        Form1.TextBox58.Clear()
        Form1.TextBox59.Clear()
        '60 to 69
        Form1.TextBox60.Clear()
        Form1.TextBox61.Clear()
        Form1.TextBox62.Clear()
        Form1.TextBox63.Clear()
        Form1.TextBox64.Clear()
        Form1.TextBox65.Clear()
        Form1.TextBox66.Clear()
        Form1.TextBox67.Clear()
        Form1.TextBox68.Clear()
        Form1.TextBox69.Clear()
        '70 to 79
        Form1.TextBox70.Clear()
        Form1.TextBox71.Clear()
        Form1.TextBox72.Clear()
        Form1.TextBox73.Clear()
        Form1.TextBox74.Clear()
        Form1.TextBox75.Clear()
        Form1.TextBox76.Clear()
        Form1.TextBox77.Clear()
        Form1.TextBox78.Clear()
        Form1.TextBox79.Clear()
        '80 and 81
        Form1.TextBox80.Clear()
        Form1.TextBox81.Clear()
    End Sub
    'This sub reads in from the text boxes
    Private Sub read_in(ByRef sudoku_box(,) As Byte)
        '1st ROW
        put_in_box(sudoku_box(0, 0), Form1.TextBox1)
        put_in_box(sudoku_box(0, 1), Form1.TextBox2)
        put_in_box(sudoku_box(0, 2), Form1.TextBox3)
        put_in_box(sudoku_box(0, 3), Form1.TextBox4)
        put_in_box(sudoku_box(0, 4), Form1.TextBox5)
        put_in_box(sudoku_box(0, 5), Form1.TextBox6)
        put_in_box(sudoku_box(0, 6), Form1.TextBox7)
        put_in_box(sudoku_box(0, 7), Form1.TextBox8)
        put_in_box(sudoku_box(0, 8), Form1.TextBox9)

        '2nd ROW
        put_in_box(sudoku_box(1, 0), Form1.TextBox10)
        put_in_box(sudoku_box(1, 1), Form1.TextBox11)
        put_in_box(sudoku_box(1, 2), Form1.TextBox12)
        put_in_box(sudoku_box(1, 3), Form1.TextBox13)
        put_in_box(sudoku_box(1, 4), Form1.TextBox14)
        put_in_box(sudoku_box(1, 5), Form1.TextBox15)
        put_in_box(sudoku_box(1, 6), Form1.TextBox16)
        put_in_box(sudoku_box(1, 7), Form1.TextBox17)
        put_in_box(sudoku_box(1, 8), Form1.TextBox18)

        '3rd ROW
        put_in_box(sudoku_box(2, 0), Form1.TextBox19)
        put_in_box(sudoku_box(2, 1), Form1.TextBox20)
        put_in_box(sudoku_box(2, 2), Form1.TextBox21)
        put_in_box(sudoku_box(2, 3), Form1.TextBox22)
        put_in_box(sudoku_box(2, 4), Form1.TextBox23)
        put_in_box(sudoku_box(2, 5), Form1.TextBox24)
        put_in_box(sudoku_box(2, 6), Form1.TextBox25)
        put_in_box(sudoku_box(2, 7), Form1.TextBox26)
        put_in_box(sudoku_box(2, 8), Form1.TextBox27)

        '4th ROW
        put_in_box(sudoku_box(3, 0), Form1.TextBox28)
        put_in_box(sudoku_box(3, 1), Form1.TextBox29)
        put_in_box(sudoku_box(3, 2), Form1.TextBox30)
        put_in_box(sudoku_box(3, 3), Form1.TextBox31)
        put_in_box(sudoku_box(3, 4), Form1.TextBox32)
        put_in_box(sudoku_box(3, 5), Form1.TextBox33)
        put_in_box(sudoku_box(3, 6), Form1.TextBox34)
        put_in_box(sudoku_box(3, 7), Form1.TextBox35)
        put_in_box(sudoku_box(3, 8), Form1.TextBox36)

        '5th ROW
        put_in_box(sudoku_box(4, 0), Form1.TextBox37)
        put_in_box(sudoku_box(4, 1), Form1.TextBox38)
        put_in_box(sudoku_box(4, 2), Form1.TextBox39)
        put_in_box(sudoku_box(4, 3), Form1.TextBox40)
        put_in_box(sudoku_box(4, 4), Form1.TextBox41)
        put_in_box(sudoku_box(4, 5), Form1.TextBox42)
        put_in_box(sudoku_box(4, 6), Form1.TextBox43)
        put_in_box(sudoku_box(4, 7), Form1.TextBox44)
        put_in_box(sudoku_box(4, 8), Form1.TextBox45)

        '6th ROW
        put_in_box(sudoku_box(5, 0), Form1.TextBox46)
        put_in_box(sudoku_box(5, 1), Form1.TextBox47)
        put_in_box(sudoku_box(5, 2), Form1.TextBox48)
        put_in_box(sudoku_box(5, 3), Form1.TextBox49)
        put_in_box(sudoku_box(5, 4), Form1.TextBox50)
        put_in_box(sudoku_box(5, 5), Form1.TextBox51)
        put_in_box(sudoku_box(5, 6), Form1.TextBox52)
        put_in_box(sudoku_box(5, 7), Form1.TextBox53)
        put_in_box(sudoku_box(5, 8), Form1.TextBox54)

        '7th ROW
        put_in_box(sudoku_box(6, 0), Form1.TextBox55)
        put_in_box(sudoku_box(6, 1), Form1.TextBox56)
        put_in_box(sudoku_box(6, 2), Form1.TextBox57)
        put_in_box(sudoku_box(6, 3), Form1.TextBox58)
        put_in_box(sudoku_box(6, 4), Form1.TextBox59)
        put_in_box(sudoku_box(6, 5), Form1.TextBox60)
        put_in_box(sudoku_box(6, 6), Form1.TextBox61)
        put_in_box(sudoku_box(6, 7), Form1.TextBox62)
        put_in_box(sudoku_box(6, 8), Form1.TextBox63)

        '8th ROW
        put_in_box(sudoku_box(7, 0), Form1.TextBox64)
        put_in_box(sudoku_box(7, 1), Form1.TextBox65)
        put_in_box(sudoku_box(7, 2), Form1.TextBox66)
        put_in_box(sudoku_box(7, 3), Form1.TextBox67)
        put_in_box(sudoku_box(7, 4), Form1.TextBox68)
        put_in_box(sudoku_box(7, 5), Form1.TextBox69)
        put_in_box(sudoku_box(7, 6), Form1.TextBox70)
        put_in_box(sudoku_box(7, 7), Form1.TextBox71)
        put_in_box(sudoku_box(7, 8), Form1.TextBox72)

        '9th ROW
        put_in_box(sudoku_box(8, 0), Form1.TextBox73)
        put_in_box(sudoku_box(8, 1), Form1.TextBox74)
        put_in_box(sudoku_box(8, 2), Form1.TextBox75)
        put_in_box(sudoku_box(8, 3), Form1.TextBox76)
        put_in_box(sudoku_box(8, 4), Form1.TextBox77)
        put_in_box(sudoku_box(8, 5), Form1.TextBox78)
        put_in_box(sudoku_box(8, 6), Form1.TextBox79)
        put_in_box(sudoku_box(8, 7), Form1.TextBox80)
        put_in_box(sudoku_box(8, 8), Form1.TextBox81)

    End Sub
End Module
