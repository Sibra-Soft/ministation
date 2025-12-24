Attribute VB_Name = "ModDigitDisplay"
Public Function GenerateSegmentDisplay(value As String, picContainer As PictureBox, imageList As imageList)
Dim ValueLength As Integer
Dim Pos As Integer
Dim I As Integer

ValueLength = Len(value)

For I = 0 To ValueLength - 1
    Dim Digit As Integer
    
    Digit = Mid(value, I + 1, 1)
    
    If Digit = 0 Then
        Digit = 1
    Else
        Digit = Digit + 1
    End If
    
    picContainer.PaintPicture imageList.ListImages(Digit).Picture, Pos, 0
    
    Pos = Pos + 117
Next
End Function
