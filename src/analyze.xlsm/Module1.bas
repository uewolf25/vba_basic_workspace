Attribute VB_Name = "Module1"
Sub last_count()
    ' �ŏI�s�A�ŏI��
    Dim global_x, global_y As Integer
    ' �c�̐�
    Dim counter, offset As Integer
    counter = 1
    offset = 5
    counter = counter + offset
    
    ' �Z����
    Dim cell As String
    ' �Z���̌r�����(���ƉE)
    Dim cell_bottom, cell_right As Border
    
    
    ' �\�܂ł̋󔒃Z���s���𑪂�B
'    Do
'        cell = cells(counter, 1).Address
'        Debug.Print (cell)
'        ' cell�ϐ����炠��Z���̉��̌r�����擾����B
'        Set cell_bottom = Range(cell).Borders(xlEdgeBottom)
'
'        ' �r������������\�̎n�܂�Ȃ̂�counter��ۑ��B
'        If cell_bottom.LineStyle <> -4142 Then
'            Debug.Print ("�r�����Ȃ��̂ŏ������I�����܂��B")
'            Exit Do
'        End If
'        counter = counter + 1
'    Loop
    

    ' ����loop
    Do
        ' �Z�������擾����
        cell = cells(counter, 1).Address
        Debug.Print (cell)
        ' cell�ϐ����炠��Z���̉��ƉE�̌r�����擾����B
        Set cell_bottom = Range(cell).Borders(xlEdgeBottom)
        Set cell_left = Range(cell).Borders(xlEdgeLeft)

        ' �����r�����Ȃ����i���ƉE�𓯎��ɒ��ׂ�j
        If cell_bottom.LineStyle = -4142 And cell_left.LineStyle = -4142 Then
            Debug.Print ("�r�����Ȃ��̂ŏ������I�����܂��B")
'            Debug.Print (counter)
            Exit Do
        End If
        
        ' �s���J�E���g
        counter = counter + 1
    Loop
    
    
    ' �J�E���g��������\�̍s���ɂ���B
    global_x = counter
    Debug.Print ("�ŏI�s��" & global_x)
    
    ' �ŏI��̎擾�B
    Dim sheet As Worksheet
    Set sheet = Worksheets(1)
    
    Debug.Print ("�ŏI���" & sheet.UsedRange.Columns(sheet.UsedRange.Columns.Count).Column)

End Sub

