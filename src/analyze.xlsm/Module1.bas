Sub get_cell_point()
    ' アルファベットを格納する
    Dim char(1 To 26) As String
    Dim alphabetNum As Long
'    set alphabets
    For alphabetNum = 1 To 26
        char(alphabetNum) = Chr(alphabetNum + 64)
        Next alphabetNum

'    With Selection
'    MsgBox "行1端からの座標は" & .Top & "ポイントです。" & vbCrLf & _
'       "A列端からの座標は" & .Left & "ポイントです。" & vbCrLf & _
'       "セル範囲の高さの座標は" & .Height & "ポイントです。" & vbCrLf & _
'       "セル範囲の幅の座標は" & .Width & "ポイントです。"
'    End With

    
    ' 表の左上の値の保存
    Dim topLeft() As Variant
    Dim last As Long
    last = 100
    Dim x As Long
    Dim y As Long
    ' Cells(縦, 横)
    Dim underLine As Border
    x = 1
    y = 1
    
    Do While y <= 26
    Debug.Print "----------ここから" & char(y) & "列です。----------"
        For x = 1 To last
            Set underLine = Cells(x, y).Borders(xlEdgeTop)
            
            If underLine.LineStyle <> -4142 Then
                Debug.Print "セル値あり→" & char(y) & x
                Call get_right_of_tabular(x, y, last, char())
                Exit Do
            
            Else
               ' GoTo ContinueX
                Debug.Print "------" & x
                
            End If
            Next x
            
        y = y + 1
    Loop
        
End Sub
' 表の右上を取得する。判別材料はセル値と上の罫線があるか。
Sub get_right_of_tabular(num As Long, alp As Long, last As Long, alphaArray() As String)
    Dim topLine As Border
    Dim y As Long
    y = alp
    Do While y <= last
    
    Debug.Print "-------------------右の最終列を確かめます。-------------------"
    Set topLine = Cells(num, y).Borders(xlEdgeTop)
        If Cells(num, y).Value = "" And topLine.LineStyle = -4142 Then
            ' 空白のあったセル値のひとつ前が最終列。
            Debug.Print "右上のセル番号→" & alphaArray(y - 1) & num
            Exit Do
        Else
            Debug.Print "--------------" & alphaArray(y)
        
        End If
        y = y + 1
    Loop
            
End Sub

' 最終行の取得。
Sub get_last_row()

    Dim leftY As Long
    Dim rightY As Long
    Dim leftLastRow As Long
    Dim rightLastRow As Long
    Dim lastRowL As Long
    Dim lastRowR As Long
    Dim last As Long
    
    ' 列
    leftY = 2
    rightY = 16
    
    
    leftLastRow = Cells(Rows.Count, leftY).row
    lastRowL = Cells(leftLastRow, leftY).End(xlUp).row
    
    rightLastRow = Cells(Rows.Count, rightY).row
    lastRowR = Cells(rightLastRow, rightY).End(xlUp).row
    
    ' 最終行多いほうを保存
    If lastRowL <= lastRowR Then
        last = lastRowR
    ElseIf lastRowL >= lastRowR Then
        last = lastRowL
    End If
    
    Debug.Print "--------------------最終行は" & last
    
End Sub

' 横方向のセルの罫線から次の罫線までのセル数を取得。
Sub get_width()
    Dim rightLine As Border
    ' 列
    Dim row As Long
    row = 5
    ' 行
    Dim col As Long
    col = 8
    
    Do While row < 100
        Debug.Print Cells(col, row).Address & vbCrLf
        Set rightLine = Cells(col, row).Borders(xlEdgeRight)
        ' 罫線があるとき
        If rightLine.LineStyle <> -4142 Then
            Debug.Print "結合セルはここまで：" & Cells(col, row).Address & vbCrLf
            Exit Do
        End If
        '１つ右へ
        row = row + 1
    Loop
    
End Sub

' 縦方向のセルの罫線から次の罫線までのセル数を取得。
Sub get_height()
    Dim bottomLine As Border
    ' 列
    Dim row As Long
    row = 5
    ' 行
    Dim col As Long
    col = 8
    
    Do While col < 100
        Debug.Print Cells(col, row).Address(True, False) & vbCrLf
        Set bottomLine = Cells(col, row).Borders(xlEdgeBottom)
        ' 罫線があるとき
        If bottomLine.LineStyle <> -4142 Then
            'bottomLine = Left(bottomLine, InStr(bottomLine, "$") - 1)
            Debug.Print "結合セルはここまで：" & Cells(col, row).Address & vbCrLf
            Exit Do
        End If
        '１つ右へ
        col = col + 1
    Loop

End Sub
' ヘッダー内の縦幅の座標を取得。
Sub get_point_row_header()
    Dim startX As Long
    Dim startY As Long
    Dim endX As Long
    Dim endY As Long
    
    startX = 2
    startY = 8
    endX = 3
    endY = 17
    
    Range(Cells(startY, startX), Cells(endY, endX)).Select
    With Selection
    Debug.Print "セル範囲の高さの座標は" & .Height & "ポイントです。" & vbCrLf & _
       "セル範囲の幅の座標は" & .Width & "ポイントです。" & vbCrLf
    End With
    
End Sub
