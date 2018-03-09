Attribute VB_Name = "Module1"
Sub Clear()

    If MsgBox("Do you want to clear all data?", vbOKCancel) = vbOK Then
        Sheet1.Range("A2:D5000").Clear
    End If

End Sub

