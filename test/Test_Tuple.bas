Attribute VB_Name = "Test_Tuple"
Option Explicit

Sub allTest()
    Call TestOfCount
End Sub

Private Sub TestOfCount()
    Debug.Print "Count test"
    assertEqual Tuple.Count, 0
    assertEqual Tuple.Create().Count, 0
    assertEqual Tuple.Create(1).Count, 1
End Sub

Private Sub TesfOfEquals()
    Debug.Print "Equals test"
    
    Dim tpl1 As Tuple
    Set tpl1 = Tuple.Create()
    
    assertTrue tpl1.Equals(tpl1)
    
End Sub


Private Sub assertEqual(v, expect, Optional msg As String)
    Debug.Print msg, v; "="; expect; "?",
    If v = expect Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub

Private Sub assertTrue(v, Optional msg As String)
    Debug.Print msg, "v is "; v,
    If v Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub
