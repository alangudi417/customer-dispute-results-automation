    ' Helper: Safe Dictionary Increment (to avoid macro crashing)
Private Sub DictAdd(ByVal dict As Object, ByVal key As String, ByVal value As Double)

    If Len(key) = 0 Then Exit Sub

    If Not dict.Exists(key) Then
        dict.Add key, value
    Else
        dict(key) = dict(key) + value
    End If

End Sub


'==========================================================================================


    ' Helper: Safe Dictionary Increment (in count)
Private Sub DictCount(ByVal dict As Object, ByVal key As String)

    If Len(key) = 0 Then Exit Sub

    If Not dict.Exists(key) Then
        dict.Add key, 1
    Else
        dict(key) = dict(key) + 1
    End If

End Sub