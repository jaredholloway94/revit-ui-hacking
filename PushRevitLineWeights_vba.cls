VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub PushLineWeights()
    Dim vals_str, command As String
    vals_str = GetVals()
    command = "python .\push_line_weights.py -vals " & vals_str
    return_val = Shell(command, 1)
End Sub

Private Function GetVals() As String
    Dim vals_list() As String
    Dim vals_str As String
    Dim cls, num, n, r, c As Integer
    cls = [vals].columns.Count
    num = [vals].Count
    ReDim vals_list(num)
    r = 0
    For Each Row In [vals].rows
        c = 0
        For Each Cell In Row.Cells
            n = (r * cls) + c
            vals_list(n) = CStr(Cell.Value)
            c = c + 1
            Next Cell
        r = r + 1
        Next Row
    GetVals = Join(vals_list, " ")
End Function
