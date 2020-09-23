Attribute VB_Name = "modScores"
Global names(0 To 2) As String, scores(0 To 2) As Integer

Public Function loadScores()
    Open App.Path & "/scores.txt" For Input As 1
    For b = 0 To 2
        Input #1, names(b), scores(b)
    Next b
    Close
End Function

Public Sub saveScores()
Open App.Path & "/scores.txt" For Output As 1
    For b = 0 To 2
        Write #1, names(b), scores(b)
    Next b
Close
End Sub

