Public Class Form1

    Const Title As String = "My Games v 0.1"



    'Variables for Tic-Tac-toe
    Dim user_turn, placed As Boolean
    Dim button_click, Buttons_filled, difficulty0 As Byte

    'Variables  for Sudoku
    Dim sudoku_box(8, 8) As Byte

    'Variables for Capture This
    Dim is_captured As Boolean



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Title
        RadioButton1.Select()
        Tic_Tac_Toe_Initialize(user_turn, placed, button_click, Buttons_filled)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Select Case TabControl1.SelectedIndex
            Case 0
                Tic_Tac_Toe(user_turn, button_click, placed, Buttons_filled, difficulty0)
            Case 1
                Sudoku(sudoku_box)
            Case 3
                Capture_this()
        End Select
    End Sub

    Private Sub Initialize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click, RestartGameToolStripMenuItem.Click, TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                Tic_Tac_Toe_Initialize(user_turn, placed, button_click, Buttons_filled)
            Case 1
                Sudoku_Initialize(sudoku_box)
            Case 2
                Maze_Initialize()
            Case 3
                Capture_Initialize(is_captured)
        End Select
    End Sub



    'Subroutines for setting variables in TIC-TAC_TOE
    'These 9 Subroutines place the number of button clicked in a variable
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        button_click = 1
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        button_click = 2
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        button_click = 3
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        button_click = 4
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        button_click = 5
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        button_click = 6
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        button_click = 7
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        button_click = 8
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        button_click = 9
    End Sub
    'These subroutines set the difficulty level variable
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        difficulty0 = 1
    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        difficulty0 = 2
    End Sub
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        difficulty0 = 3
    End Sub
    'Subroutines for setting variables in TIC-TAC_TOE END



    'Subroutines For the MAZE
    Public Sub Maze_Initialize()
        Timer1.Stop()
        Dim startingpoint As Point

        startingpoint = Me.Location + TabControl1.Location + TabControl1.Location + TabControl1.Location
        startingpoint.Offset(15, 15)
        Cursor.Position = startingpoint
    End Sub
    Private Sub maze_done(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles finish.MouseEnter
        'For winning the maze game
        MsgBox("Congratulations, you completed the maze")
    End Sub
    Private Sub wall_hit(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.MouseEnter, Label4.MouseEnter, Label3.MouseEnter, Label2.MouseEnter, Label9.MouseEnter, Label8.MouseEnter, Label7.MouseEnter, Label6.MouseEnter, Label17.MouseEnter, Label16.MouseEnter, Label15.MouseEnter, Label14.MouseEnter, Label12.MouseEnter, Label11.MouseEnter, Label21.MouseEnter, Label20.MouseEnter, Label19.MouseEnter, Label18.MouseEnter, Label28.MouseEnter, Label27.MouseEnter, Label26.MouseEnter, Label25.MouseEnter, Label24.MouseEnter, Label23.MouseEnter, Label22.MouseEnter, Label30.MouseEnter, Label29.MouseEnter, Label13.MouseEnter, Label10.MouseEnter, Label32.MouseEnter, Label31.MouseEnter, Label55.MouseEnter, Label54.MouseEnter, Label53.MouseEnter, Label52.MouseEnter, Label51.MouseEnter, Label50.MouseEnter, Label49.MouseEnter, Label48.MouseEnter, Label47.MouseEnter, Label46.MouseEnter, Label45.MouseEnter, Label44.MouseEnter, Label43.MouseEnter, Label42.MouseEnter, Label41.MouseEnter, Label40.MouseEnter, Label39.MouseEnter, Label38.MouseEnter, Label37.MouseEnter, Label35.MouseEnter, Label34.MouseEnter, Label33.MouseEnter, Label90.MouseEnter, Label89.MouseEnter, Label88.MouseEnter, Label87.MouseEnter, Label86.MouseEnter, Label85.MouseEnter, Label84.MouseEnter, Label83.MouseEnter, Label82.MouseEnter, Label81.MouseEnter, Label80.MouseEnter, Label79.MouseEnter, Label78.MouseEnter, Label77.MouseEnter, Label76.MouseEnter, Label75.MouseEnter, Label74.MouseEnter, Label73.MouseEnter, Label72.MouseEnter, Label71.MouseEnter, Label70.MouseEnter, Label69.MouseEnter, Label68.MouseEnter, Label67.MouseEnter, Label66.MouseEnter, Label65.MouseEnter, Label64.MouseEnter, Label63.MouseEnter, Label62.MouseEnter, Label61.MouseEnter, Label60.MouseEnter, Label59.MouseEnter, Label58.MouseEnter, Label57.MouseEnter, Label56.MouseEnter, Label36.MouseEnter, Label99.MouseEnter, Label98.MouseEnter, Label97.MouseEnter, Label96.MouseEnter, Label95.MouseEnter, Label94.MouseEnter, Label93.MouseEnter, Label92.MouseEnter, Label91.MouseEnter, Label121.MouseEnter, Label120.MouseEnter, Label119.MouseEnter, Label118.MouseEnter, Label117.MouseEnter, Label116.MouseEnter, Label115.MouseEnter, Label114.MouseEnter, Label113.MouseEnter, Label112.MouseEnter, Label111.MouseEnter, Label110.MouseEnter, Label109.MouseEnter, Label108.MouseEnter, Label107.MouseEnter, Label106.MouseEnter, Label105.MouseEnter, Label104.MouseEnter, Label103.MouseEnter, Label102.MouseEnter, Label101.MouseEnter, Label100.MouseEnter, TabPage3.MouseEnter, Label146.MouseEnter, Label145.MouseEnter, Label144.MouseEnter, Label143.MouseEnter, Label142.MouseEnter, Label141.MouseEnter, Label140.MouseEnter, Label139.MouseEnter, Label138.MouseEnter, Label137.MouseEnter, Label136.MouseEnter, Label135.MouseEnter, Label134.MouseEnter, Label133.MouseEnter, Label132.MouseEnter, Label131.MouseEnter, Label130.MouseEnter, Label129.MouseEnter, Label128.MouseEnter, Label127.MouseEnter, Label126.MouseEnter, Label125.MouseEnter, Label124.MouseEnter, Label123.MouseEnter, Label122.MouseEnter
        Maze_Initialize()
    End Sub
    Private Sub ExitMazeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitMazeToolStripMenuItem.Click
        Cursor.Position = Me.Location
    End Sub
    'Subroutines For the MAZE END



    'Subroutines For CAPTURE THIS
    Private Sub captured(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cap.MouseEnter
        If (is_captured = False) Then
            Timer1.Stop()
            MsgBox("Congrats. You captured it")
            is_captured = True
        End If
    End Sub
    Private Sub capture_difficulty_changed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged, RadioButton5.CheckedChanged, RadioButton4.CheckedChanged
        difficulty_select()
    End Sub
    
End Class
