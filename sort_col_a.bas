Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Cells.Column <> 1 Then Exit Sub
    If Target.Cells.Column = 1 Then
        Columns("A:A").Sort Key1:=Range("A1"), _
                                  Order1:=xlAscending, _
                                  Header:=xlGuess, _
                                  OrderCustom:=1, _
                                  MatchCase:=False, _
                                  Orientation:=xlTopToBottom, _
                                  DataOption1:=xlSortNormal
    End If
End Sub