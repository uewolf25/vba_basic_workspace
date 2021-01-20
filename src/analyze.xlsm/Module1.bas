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
    
    For y = 1 To 26
    Debug.Print "----------ここから" & char(y) & "列です。----------"
        For x = 1 To last
            Set underLine = Cells(x, y).Borders(xlEdgeTop)
            
            If underLine.LineStyle = 1 Then
                Debug.Print "左上のセル番号→" & char(y) & x
                Call get_right_of_tabular(x, y, last, char())
                Exit Sub
            
            Else
               ' GoTo ContinueX
                Debug.Print "------" & x
                
            End If
            Next x
            
        Next y
        
' 疑似continue文
'ContinueX:
'    Next x

        
End Sub
' 表の右上を取得する。判別材料はセル値と上の罫線があるか。
Sub get_right_of_tabular(num As Long, alp As Long, last As Long, alphaArray() As String)
    Dim topLine As Border
    
    
    For y = alp To last
    Debug.Print "-------------------右の最終列を確かめます。-------------------"
    Set topLine = Cells(num, y).Borders(xlEdgeTop)
        If Cells(num, y).Value = "" And topLine.LineStyle = -4142 Then
            ' 空白のあったセル値のひとつ前が最終列。
            Debug.Print "右上のセル番号→" & alphaArray(y - 1) & num
            Exit Sub
        Else
            Debug.Print "--------------" & alphaArray(y)
        
        End If
        Next y
            
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
    
    
    leftLastRow = Cells(Rows.Count, leftY).Row
    lastRowL = Cells(leftLastRow, leftY).End(xlUp).Row
    
    rightLastRow = Cells(Rows.Count, rightY).Row
    lastRowR = Cells(rightLastRow, rightY).End(xlUp).Row
    
    ' 最終行多いほうを保存
    If lastRowL <= lastRowR Then
        last = lastRowR
    ElseIf lastRowL >= lastRowR Then
        last = lastRowL
    End If
    
    Debug.Print "--------------------最終行は" & last
    
End Sub




