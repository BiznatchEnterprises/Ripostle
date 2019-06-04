Attribute VB_Name = "Module1"
Public Function FormatFileSize(Size As Long) As String
    Dim strSize As String, strSize2 As String, strSize3 As String, strSize4 As String
    Dim FormattedSize As String
    Dim i As Integer
    strSize = Str(Size)
Select Case Len(strSize)
        Case 5
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 2)
        strSize3 = strSize3 + strSize2 + " Kb"
        Case 6
       strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 3)
        strSize3 = strSize3 + strSize2 + " Kb"
        Case 7
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 4)
        strSize3 = strSize3 + strSize2 + " Kb"
        Case 8
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 5)
        strSize3 = strSize3 + strSize2
        
        strSize4 = Right(strSize3, 7)
        strSize4 = "," + strSize4
        strSize3 = Left(strSize, 2)
        strSize3 = strSize3 + strSize4 + " Mb"
        Case 9

        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 5)
        strSize3 = strSize3 + strSize2
        
        strSize4 = Right(strSize3, 7)
        strSize4 = "," + strSize4
        strSize3 = Left(strSize, 3)
        strSize3 = strSize3 + strSize4 + " Mb"
        Case Else
        strSize3 = strSize + " bytes"
    End Select



FormattedSize = strSize3


    FormatFileSize = FormattedSize
    End Function


