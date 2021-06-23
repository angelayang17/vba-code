Attribute VB_Name = "Module3"
Sub Button2_Click()
    Range("AK10").Select 'Select cell AK10
    
    Do Until Selection.Interior.Color = vbRed
    
        If Selection.Offset(-1, 0).Interior.Color = vbWhite Or Selection.Offset(-1, 0).Interior.Color = vbRed Then
            'move up
            Selection.Interior.Color = vbYellow
            Selection.Offset(-1, 0).Select
        ElseIf Selection.Offset(0, 1).Interior.Color = vbWhite Or Selection.Offset(0, 1).Interior.Color = vbRed Then
            'move right
            Selection.Interior.Color = vbYellow
            Selection.Offset(0, 1).Select
        ElseIf Selection.Offset(1, 0).Interior.Color = vbWhite Or Selection.Offset(1, 0).Interior.Color = vbRed Then
            'move down
            Selection.Interior.Color = vbYellow
            Selection.Offset(1, 0).Select
        ElseIf Selection.Offset(0, -1).Interior.Color = vbWhite Or Selection.Offset(0, -1).Interior.Color = vbRed Then
            'move left
            Selection.Interior.Color = vbYellow
            Selection.Offset(0, -1).Select
            
        'Backtrack (backtracked cells are cyan)
        ElseIf Selection.Offset(1, 0).Interior.Color = vbBlack Or Selection.Offset(-1, 0).Interior.Color = vbBlack Or Selection.Offset(0, 1).Interior.Color = vbBlack Or Selection.Offset(0, -1).Interior.Color = vbBlack Then
                
                Do Until Selection.Offset(1, 0).Interior.Color = vbWhite Or Selection.Offset(-1, 0).Interior.Color = vbWhite Or Selection.Offset(0, 1).Interior.Color = vbWhite Or Selection.Offset(0, -1).Interior.Color = vbWhite
                    If Selection.Offset(-1, 0).Interior.Color = vbYellow Then
                        'move up
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(-1, 0).Select
                    ElseIf Selection.Offset(0, 1).Interior.Color = vbYellow Then
                        'move right
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(0, 1).Select
                    ElseIf Selection.Offset(1, 0).Interior.Color = vbYellow Then
                        'move down
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(1, 0).Select
                    ElseIf Selection.Offset(0, -1).Interior.Color = vbYellow Then
                        'move left
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(0, -1).Select
                    End If
                Loop

        Else
            'Show a message that no other possible moves are possible
            MsgBox "There are no possible moves"
            Exit Do 'This stops the loop entirely
        End If
        
    Loop

End Sub
Sub Button3_Click()
    Range("AK10").Select 'Select cell AK10
    
    Do Until Selection.Interior.Color = vbRed
    
        If Selection.Offset(-1, 0).Interior.Color = vbWhite Or Selection.Offset(-1, 0).Interior.Color = vbRed Then
            'move up
            Selection.Interior.Color = vbYellow
            Selection.Offset(-1, 0).Select
        ElseIf Selection.Offset(0, 1).Interior.Color = vbWhite Or Selection.Offset(0, 1).Interior.Color = vbRed Then
            'move right
            Selection.Interior.Color = vbYellow
            Selection.Offset(0, 1).Select
        ElseIf Selection.Offset(1, 0).Interior.Color = vbWhite Or Selection.Offset(1, 0).Interior.Color = vbRed Then
            'move down
            Selection.Interior.Color = vbYellow
            Selection.Offset(1, 0).Select
        ElseIf Selection.Offset(0, -1).Interior.Color = vbWhite Or Selection.Offset(0, -1).Interior.Color = vbRed Then
            'move left
            Selection.Interior.Color = vbYellow
            Selection.Offset(0, -1).Select
            
        'Backtrack (backtracked cells are cyan)
        ElseIf Selection.Offset(1, 0).Interior.Color = vbBlack Or Selection.Offset(-1, 0).Interior.Color = vbBlack Or Selection.Offset(0, 1).Interior.Color = vbBlack Or Selection.Offset(0, -1).Interior.Color = vbBlack Then
                
                Do Until Selection.Offset(1, 0).Interior.Color = vbWhite Or Selection.Offset(-1, 0).Interior.Color = vbWhite Or Selection.Offset(0, 1).Interior.Color = vbWhite Or Selection.Offset(0, -1).Interior.Color = vbWhite
                    If Selection.Offset(-1, 0).Interior.Color = vbYellow Then
                        'move up
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(-1, 0).Select
                    ElseIf Selection.Offset(0, 1).Interior.Color = vbYellow Then
                        'move right
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(0, 1).Select
                    ElseIf Selection.Offset(1, 0).Interior.Color = vbYellow Then
                        'move down
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(1, 0).Select
                    ElseIf Selection.Offset(0, -1).Interior.Color = vbYellow Then
                        'move left
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(0, -1).Select
                    End If
                
                Loop
                
                Application.Wait (Now + 0.00001)

        Else
            'Show a message that no other possible moves are possible
            MsgBox "There are no possible moves"
            Exit Do 'This stops the loop entirely
        End If

    Loop
    
    Application.Wait (Now + 0.00001)
End Sub
Sub Button11_Click()
    Range("AK10").Select 'Select cell AK10
    
    Do Until Selection.Interior.Color = vbRed
    
        If Selection.Offset(-1, 0).Interior.Color = vbWhite Or Selection.Offset(-1, 0).Interior.Color = vbRed Then
            'move up
            Selection.Interior.Color = vbYellow
            Selection.Offset(-1, 0).Select
        ElseIf Selection.Offset(0, 1).Interior.Color = vbWhite Or Selection.Offset(0, 1).Interior.Color = vbRed Then
            'move right
            Selection.Interior.Color = vbYellow
            Selection.Offset(0, 1).Select
        ElseIf Selection.Offset(1, 0).Interior.Color = vbWhite Or Selection.Offset(1, 0).Interior.Color = vbRed Then
            'move down
            Selection.Interior.Color = vbYellow
            Selection.Offset(1, 0).Select
        ElseIf Selection.Offset(0, -1).Interior.Color = vbWhite Or Selection.Offset(0, -1).Interior.Color = vbRed Then
            'move left
            Selection.Interior.Color = vbYellow
            Selection.Offset(0, -1).Select
            
        'Backtrack (backtracked cells are cyan)
        ElseIf Selection.Offset(1, 0).Interior.Color = vbBlack Or Selection.Offset(-1, 0).Interior.Color = vbBlack Or Selection.Offset(0, 1).Interior.Color = vbBlack Or Selection.Offset(0, -1).Interior.Color = vbBlack Then
                
                Do Until Selection.Offset(1, 0).Interior.Color = vbWhite Or Selection.Offset(-1, 0).Interior.Color = vbWhite Or Selection.Offset(0, 1).Interior.Color = vbWhite Or Selection.Offset(0, -1).Interior.Color = vbWhite
                    If Selection.Offset(-1, 0).Interior.Color = vbYellow Then
                        'move up
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(-1, 0).Select
                    ElseIf Selection.Offset(0, 1).Interior.Color = vbYellow Then
                        'move right
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(0, 1).Select
                    ElseIf Selection.Offset(1, 0).Interior.Color = vbYellow Then
                        'move down
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(1, 0).Select
                    ElseIf Selection.Offset(0, -1).Interior.Color = vbYellow Then
                        'move left
                        Selection.Interior.Color = vbCyan
                        Selection.Offset(0, -1).Select
                    End If
                Loop

        Else
            'Show a message that no other possible moves are possible
            MsgBox "There are no possible moves"
            Exit Do 'This stops the loop entirely
        End If
        
    Loop
End Sub
