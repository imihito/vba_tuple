VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'version 0.0.1

Option Explicit

'(0 To n)
Private items_() As Variant

'Constructor
'`Tuple.Create()` usage is not thread safe (but may be vba is single thread).
Public Function Create(ByRef iItems() As Variant) As Tuple
    If Not isInitialized Then Call VBA.Err().Raise(5)
    If Me Is Tuple Then Set Tuple = Nothing
    
    Let items_() = iItems()
    Set Create = Me
End Function

'is inteanl array initialized?
Private Property Get isInitialized() As Boolean
    Let isInitialized = CBool(Not Not items_())
End Property

'
Public Function Assign(ParamArray assignVariables() As Variant) As Tuple
    Set Assign = Me
    
    Dim i As Long
    For i = LBound(assignVariables()) To UBound(assignVariables())
        setValue assignVariables(i), items_(i)
    Next i
    
End Function

'Index access. 0 start.
Public Property Get Item(Index As Long) As Variant
    Call setValue(Item, items_(Index))
End Property

Public Property Get Count() As Long
    If isInitialized Then Let Count = arraySizeOf(items_)
End Property


Public Function ToArray() As Variant()
    Let ToArray = items_()
End Function


Public Function Equals( _
            t As Tuple, _
            Optional Compare As VBA.VbCompareMethod = VBA.VbCompareMethod.vbBinaryCompare _
        ) As Boolean
    
    Err.Raise 5, , "not implemented"
    Let Equals = internalArrayEquals(Me.ToArray(), t.ToArray())
End Function

'common internal functions
Private Sub setValue(ByRef outVariable As Variant, ByRef setVal As Variant)
    If VBA.IsObject(setVal) Then _
        Set outVariable = setVal _
    Else _
        Let outVariable = setVal
End Sub

Private Function arraySizeOf( _
            iArray As Variant, _
            Optional iDimension As Long = 1 _
        ) As Long
        
    If Not VBA.IsArray(iArray) Then Call VBA.Err().Raise(13)
    
    On Error Resume Next
        Let arraySizeOf = (UBound(iArray, iDimension) - LBound(iArray, iDimension) + 1)
    On Error GoTo 0
End Function

Private Function arrayDimensions(iArray As Variant) As Long
    If Not VBA.IsArray(iArray) Then Call VBA.Err().Raise(13)
    
    Const VBA_ArrayDimMax = 64
    On Error GoTo Last
        Dim d As Long
        For d = 1 To VBA_ArrayDimMax
            Dim no_mean_var As Long
            no_mean_var = LBound(iArray, d)
        Next d
    On Error GoTo 0
Last:
    Let arrayDimensions = d - 1
End Function


'for Equals Method
Private Function internalArrayEquals(arr1 As Variant, arr2 As Variant) As Boolean
    ' not imolemented
    On Error GoTo FAIL
    Dim i As Long
    For i = 0 To Me.Count - 1
        
    Next i
FAIL:
End Function

Private Function valueEquals(a As Variant, b As Variant, Compare As VBA.VbCompareMethod) As Boolean
    Dim a_isObj As Boolean, b_isObj As Boolean
    a_isObj = VBA.IsObject(a)
    b_isObj = VBA.IsObject(b)
    If a_isObj Or b_isObj Then
        If a_isObj And b_isObj Then
            Let valueEquals = (a Is b)
        Else
            Let valueEquals = False
        End If
        Exit Function
    End If
    
    Dim a_isArr As Boolean, b_isArr As Boolean
    a_isArr = VBA.IsArray(a)
    b_isArr = VBA.IsArray(b)
    If a_isArr Or b_isArr Then
        If a_isArr And b_isArr Then
            Let valueEquals = internalArrayEquals(a, b)
        Else
            Let valueEquals = False
        End If
        Exit Function
    End If
    
    If VBA.VarType(a) = vbString Or VBA.VarType(b) = vbString Then
        Let valueEquals = (VBA.StrComp(a, b, Compare) = 0)
    Else
        'Expect both type is numeric type.
        Let valueEquals = (a = b)
    End If
    
End Function
