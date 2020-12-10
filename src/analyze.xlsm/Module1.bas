Attribute VB_Name = "Module1"
Sub last_count()
    ' 最終行、最終列
    Dim global_x, global_y As Integer
    ' 縦の数
    Dim counter, offset As Integer
    counter = 1
    offset = 5
    counter = counter + offset
    
    ' セル名
    Dim cell As String
    ' セルの罫線情報(下と右)
    Dim cell_bottom, cell_right As Border
    
    
    ' 表までの空白セル行数を測る。
'    Do
'        cell = cells(counter, 1).Address
'        Debug.Print (cell)
'        ' cell変数からあるセルの下の罫線を取得する。
'        Set cell_bottom = Range(cell).Borders(xlEdgeBottom)
'
'        ' 罫線が見つけたら表の始まりなのでcounterを保存。
'        If cell_bottom.LineStyle <> -4142 Then
'            Debug.Print ("罫線がないので処理を終了します。")
'            Exit Do
'        End If
'        counter = counter + 1
'    Loop
    

    ' 無限loop
    Do
        ' セル名を取得する
        cell = cells(counter, 1).Address
        Debug.Print (cell)
        ' cell変数からあるセルの下と右の罫線を取得する。
        Set cell_bottom = Range(cell).Borders(xlEdgeBottom)
        Set cell_left = Range(cell).Borders(xlEdgeLeft)

        ' もし罫線がない時（下と右を同時に調べる）
        If cell_bottom.LineStyle = -4142 And cell_left.LineStyle = -4142 Then
            Debug.Print ("罫線がないので処理を終了します。")
'            Debug.Print (counter)
            Exit Do
        End If
        
        ' 行をカウント
        counter = counter + 1
    Loop
    
    
    ' カウントした数を表の行数にする。
    global_x = counter
    Debug.Print ("最終行は" & global_x)
    
    ' 最終列の取得。
    Dim sheet As Worksheet
    Set sheet = Worksheets(1)
    
    Debug.Print ("最終列は" & sheet.UsedRange.Columns(sheet.UsedRange.Columns.Count).Column)

End Sub

