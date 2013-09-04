Module Subs_Tic_Tac_Toe

    'PLACES SIGNS ON BUTTONS in Tic-Tac-Toe
    Private Sub button_click_in_tic_tac_toe(ByVal button_click As Byte, ByVal user_turn As Boolean, ByRef placed As Boolean)

        Dim Sign As String = ""

        If user_turn = True Then
            Sign = "X"
        Else
            Sign = "O"
        End If

        Select Case button_click
            Case 1
                If Form1.Button1.Text = "" Then
                    Form1.Button1.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 2
                If Form1.Button2.Text = "" Then
                    Form1.Button2.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 3
                If Form1.Button3.Text = "" Then
                    Form1.Button3.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 4
                If Form1.Button4.Text = "" Then
                    Form1.Button4.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 5
                If Form1.Button5.Text = "" Then
                    Form1.Button5.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 6
                If Form1.Button6.Text = "" Then
                    Form1.Button6.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 7
                If Form1.Button7.Text = "" Then
                    Form1.Button7.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 8
                If Form1.Button8.Text = "" Then
                    Form1.Button8.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
            Case 9
                If Form1.Button9.Text = "" Then
                    Form1.Button9.Text = Sign
                    placed = True
                Else
                    placed = False
                End If
        End Select
    End Sub
    'INITIALIZES Tic-Tac-Toe
    Public Sub Tic_Tac_Toe_Initialize(ByRef user_turn As Boolean, ByRef placed As Boolean, ByRef button_click As Byte, ByRef buttons_filled As Byte)

        Form1.Timer1.Stop()

        user_turn = True
        placed = False
        button_click = 0
        buttons_filled = 0

        Form1.Label1.Text = "User's Turn"
        Form1.Timer1.Interval = 1
        Form1.Timer1.Start()

        Form1.Button1.Text = ""
        Form1.Button2.Text = ""
        Form1.Button3.Text = ""
        Form1.Button4.Text = ""
        Form1.Button5.Text = ""
        Form1.Button6.Text = ""
        Form1.Button7.Text = ""
        Form1.Button8.Text = ""
        Form1.Button9.Text = ""
    End Sub
    'ACTUAL GAME of Tic-Tac-Toe
    Public Sub Tic_Tac_Toe(ByRef user_turn As Boolean, ByRef button_click As Byte, ByRef placed As Boolean, ByRef Buttons_filled As Byte, ByVal difficulty As Byte)

        If ((user_turn = True) And (button_click <> 0)) Then 'This If is for user's turn

            button_click_in_tic_tac_toe(button_click, user_turn, placed)

            If (placed = True) Then
                user_turn = False
                placed = False
                Form1.Label1.Text = "Computer's Turn"
                Buttons_filled += 1
            End If

        Else 'This else is for computer's turn

            If (Buttons_filled <> 0) Then
                Do 'This loop decides at which place computer will place the sign

                    If (difficulty = 1) Then 'This is for EASY difficulty
                        button_click = Int(Rnd() * 9) + 1
                    ElseIf (difficulty = 2) Then 'This is for NORMAL difficulty
                        If (Form1.Button5.Text = "") Then
                            button_click = 5
                        ElseIf ((Form1.Button2.Text = "X") Or (Form1.Button4.Text = "X")) And (Form1.Button1.Text = "") Then
                            button_click = 1
                        ElseIf ((Form1.Button2.Text = "X") Or (Form1.Button6.Text = "X")) And (Form1.Button3.Text = "") Then
                            button_click = 3
                        ElseIf ((Form1.Button8.Text = "X") Or (Form1.Button6.Text = "X")) And (Form1.Button9.Text = "") Then
                            button_click = 9
                        ElseIf ((Form1.Button8.Text = "X") Or (Form1.Button4.Text = "X")) And (Form1.Button7.Text = "") Then
                            button_click = 7
                        Else
                            button_click = Int(Rnd() * 9) + 1
                        End If
                    Else 'THIS IS IN CASE THERE IS SOME OTHER DIFFICULTY
                        button_click = Int(Rnd() * 9) + 1
                    End If

                    button_click_in_tic_tac_toe(button_click, user_turn, placed)
                Loop Until placed = True


                Buttons_filled += 1
                user_turn = True
                placed = False
                Form1.Label1.Text = "User's Turn"
            End If

        End If 'Computer's turn ends here

        Check_Tic_Tac_Toe(Buttons_filled)
    End Sub



    'Subroutines for finding and displaying who wins
    Private Sub Check_Tic_Tac_Toe(ByVal Buttons_filled As Byte)
        If (Form1.Button5.Text = Form1.Button1.Text) And (Form1.Button5.Text = Form1.Button9.Text) And (Form1.Button5.Text <> "") Then
            Display1() '1st diagonal
        ElseIf (Form1.Button5.Text = Form1.Button3.Text) And (Form1.Button5.Text = Form1.Button7.Text) And (Form1.Button5.Text <> "") Then
            Display1() '2nd diagonal
        ElseIf (Form1.Button5.Text = Form1.Button2.Text) And (Form1.Button5.Text = Form1.Button8.Text) And (Form1.Button5.Text <> "") Then
            Display1() '2nd column
        ElseIf (Form1.Button5.Text = Form1.Button4.Text) And (Form1.Button5.Text = Form1.Button6.Text) And (Form1.Button5.Text <> "") Then
            Display1() '2nd row

        ElseIf (Form1.Button7.Text = Form1.Button1.Text) And (Form1.Button7.Text = Form1.Button4.Text) And (Form1.Button7.Text <> "") Then
            Display2() '1st column
        ElseIf (Form1.Button7.Text = Form1.Button8.Text) And (Form1.Button7.Text = Form1.Button9.Text) And (Form1.Button7.Text <> "") Then
            Display2() '3rd row

        ElseIf (Form1.Button3.Text = Form1.Button1.Text) And (Form1.Button3.Text = Form1.Button2.Text) And (Form1.Button3.Text <> "") Then
            Display3() '1st row
        ElseIf (Form1.Button3.Text = Form1.Button6.Text) And (Form1.Button3.Text = Form1.Button9.Text) And (Form1.Button3.Text <> "") Then
            Display3() '3rd column

        ElseIf Buttons_filled = 9 Then
            Form1.Timer1.Stop()
            Form1.Label1.Text = "Draw"
        End If
    End Sub
    Private Sub Display1()
        Form1.Timer1.Stop()
        If (Form1.Button5.Text = "X") Then
            Form1.Label1.Text = "User Wins"
        Else
            Form1.Label1.Text = "Computer Wins"
        End If
    End Sub
    Private Sub Display2()
        Form1.Timer1.Stop()
        If (Form1.Button7.Text = "X") Then
            Form1.Label1.Text = "User Wins"
        Else
            Form1.Label1.Text = "Computer Wins"
        End If
    End Sub
    Private Sub Display3()
        Form1.Timer1.Stop()
        If (Form1.Button3.Text = "X") Then
            Form1.Label1.Text = "User Wins"
        Else
            Form1.Label1.Text = "Computer Wins"
        End If
    End Sub
    'Subroutines for finding and displaying who wins END


End Module
