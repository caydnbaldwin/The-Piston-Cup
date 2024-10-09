Option Explicit

Sub Kachow()

' assign variables
Dim lane_1 As Integer
Dim lane_2 As Integer
Dim lane_3 As Integer
Dim random_number As Integer

' clear previous values
Cells(7, 1).Value = ""
Cells(7, 1).Interior.Color = RGB(255, 255, 255)
Range("B3:H5").Interior.Color = RGB(255, 255, 255)
Application.Wait (Now + TimeValue("00:00:01"))

' set starting values
lane_1 = 1
lane_2 = 1
lane_3 = 1

' loop while lanes have not reached 8
While lane_1 < 8 And lane_2 < 8 And lane_3 < 8

' Check if a lane is 3 or more cells ahead or behind

    ' lane 1 is too far ahead of lane 2
    If lane_1 - 2 > lane_2 Then
            random_number = 2
            
        ' lane 1 is too far ahead of lane 3
        ElseIf lane_1 - 2 > lane_3 Then
            random_number = 3
            
        ' lane 2 is too far ahead of lane 1
        ElseIf lane_2 - 2 > lane_1 Then
            random_number = 1
            
        ' lane 2 is too far ahead of lane 3
        ElseIf lane_2 - 2 > lane_3 Then
            random_number = 3
        
        ' lane 3 is too far ahead of lane 1
        ElseIf lane_3 - 2 > lane_1 Then
            random_number = 1
            
        ' lane 3 is too far ahead of lane 2
        ElseIf lane_3 - 2 > lane_2 Then
            random_number = 2
            
        ' if nobody is too far ahead, proceed as normal
        Else
            random_number = Application.WorksheetFunction.RandBetween(1, 3)
            
    End If
    
' Increment the count of the lane that the random number chose, then paint the new cell that belongs to that lane
    
    ' if random number is 1, paint lane 1
    If random_number = 1 Then
            lane_1 = lane_1 + 1
            Cells(3, lane_1).Interior.Color = RGB(255, 0, 0)
                
        ' if random number is 2, paint lane 2
        ElseIf random_number = 2 Then
            lane_2 = lane_2 + 1
            Cells(4, lane_2).Interior.Color = RGB(0, 255, 0)
            
        ' if random number is 3, paint lane 3
        ElseIf random_number = 3 Then
            lane_3 = lane_3 + 1
            Cells(5, lane_3).Interior.Color = RGB(0, 255, 255)
            
    End If
    
    ' time buffer for suspense
    Application.Wait (Now + TimeValue("00:00:01"))
    
Wend

' print winner

' if lane 1 won
If lane_1 = 8 Then
        Cells(7, 1).Interior.Color = RGB(255, 0, 0)
        Cells(7, 1).Font.Color = RGB(255, 255, 0)
        Cells(7, 1).Value = "The winner is Lightning McQueen!"
    
    ' if lane 2 won
    ElseIf lane_2 = 8 Then
        Cells(7, 1).Interior.Color = RGB(0, 255, 0)
        Cells(7, 1).Font.Color = RGB(0, 0, 0)
        Cells(7, 1).Value = "The winner is Chick Hicks!"
        
    ' if lane 3 won
    ElseIf lane_3 = 8 Then
        Cells(7, 1).Interior.Color = RGB(0, 255, 255)
        Cells(7, 1).Font.Color = RGB(255, 255, 255)
        Cells(7, 1).Value = "The winner is The King!"

End If
        
Beep
End Sub
