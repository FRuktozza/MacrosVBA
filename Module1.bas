Sub ЖёлтыйЦвет()
    For Each cell In Selection
        cell.Interior.Color = vbYellow
    Next
End Sub

Function МОДУЛЬРАЗНИЦЫ(a, b)
    МОДУЛЬРАЗНИЦЫ = Abs(a - b)
End Function
