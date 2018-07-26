Attribute VB_Name = "Sample_Tuple"
Option Explicit

Sub Sample()
    
'create
    Dim tpl1 As Tuple
    Set tpl1 = Tuple.Create(1, "a", Now) 'not thread safe.
    
    Dim tpl2 As Tuple
    Set tpl2 = New Tuple
    Set tpl2 = tpl2.Create(1, "a", Now)
    
    
'assign
    Dim n As Long, a As String, d As Date
    tpl1.Assign n, a, d 'n = 1, a = "a", d = Now
    Stop
    
End Sub

