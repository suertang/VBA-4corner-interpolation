Public Function inter2d(a, b, c, d, newx)
If newx < a Then
    inter2d = c
ElseIf newx > b Then
    inter2d = d
Else
    k = (c - d) / (a - b)
    bias = c - k * a
    inter2d = k * newx + bias
End If
End Function
Function inter(x, y, col, idx, drng As Range)

Dim cols, rows
cols = Application.Index(col.Value, 1, 0)
rows = Application.Transpose(Application.Index(idx.Value, 0, 1))
d = drng.Value


x_o = 0
y_o = 0
For i = 1 To UBound(cols)
    If cols(1) > x Then
        x_o = 0
        Exit For
    ElseIf cols(i) > x Then
        x_o = i
        Exit For
    Else
        x_o = i
    End If
Next
For i = 1 To UBound(rows)
    If rows(1) > y Then
        y_o = 0
        Exit For
    ElseIf rows(i) > y Then
        y_o = i
        Exit For
    Else
        y_o = i
    End If
Next
'Debug.Print x_o, y_o

If x_o = 0 Then
    intx1 = d(y_o - 1, 1)
    intx2 = d(y_o, 1)
Else
    intx1 = inter2d(cols(x_o - 1), cols(x_o), d(y_o - 1, x_o - 1), d(y_o - 1, x_o), x)
    intx2 = inter2d(cols(x_o - 1), cols(x_o), d(y_o, x_o - 1), d(y_o, x_o), x)
End If

If y_o = 0 Then
    inty = inter2d(rows(1), rows(1), intx1, intx2, y)
Else
    inty = inter2d(rows(y_o - 1), rows(y_o), intx1, intx2, y)
End If

Debug.Print inty
inter = inty

End Function

