我重新做了一下你的功能，不知道是否合用，你把代码贴到sheet1中，代码如下：
1、 在工作表sheet1上放置一个listbox和一个textbox，设置它们的Visiable属性为False。
2、 在sheet2的A列上输入一些人名，例如：   张三、张三丰、小明、小李飞刀……最好是有同姓的
3、 用VBA控制，当活动单元格为sheet1 A列中的单元格时，listbox和textbox的visiable属性都设为True，并且设置他们的left、top、width、height属性
4、 将sheet2表中A列上所有的值填到listbox中
到这里这个功能就已经基本实现了，剩下的工作就是做得更人性化一些，更好用一些，所有代码如下：


Private Sub infoList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.infoList.listCount > 0 Then
        Selection = Me.infoList.Value
    End If
'**将活动单元格切换到同列的下一个
    Sheet1.Cells(Selection.Row + 1, Selection.Column).Select
End Sub

Private Sub txtContext_Change()
    Dim arrList() As String
    
    If Me.infoList.listCount > 0 Then
        ReDim arrList(Me.infoList.listCount - 1)
    Else
        Exit Sub
    End If
    
    j% = 0
    For i% = 0 To UBound(arrList)
        If Me.txtContext.Text = Mid(Me.infoList.List(i), 1, Len(Me.txtContext.Text)) Then
            arrList(j%) = Me.infoList.List(i%)
            j% = j% + 1
        End If
    Next
    If arrList(0) <> "" Then
        '**若有符合条件的内容则，列表框清空后重新填入符合条件的内容
        Me.infoList.Clear
        For i% = 0 To UBound(arrList)
            If arrList(i%) <> "" Then
                Me.infoList.AddItem arrList(i%)
            End If
        Next
    End If
End Sub

Private Sub txtContext_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '**判断是否按下回车
    If KeyCode = 13 Then
        '**将文本框的内容填入当前活动单元格
        Selection = Me.txtContext.Text
        '**将活动单元格切换到同列的下一个
        Sheet1.Cells(Selection.Row + 1, Selection.Column).Select
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    '**判断用户是否选中第一列，只有当前活动单元格为第一列中的单元格时才显示列表
    If Target.Column = 1 Then
        '**设置textBox
        With Me.txtContext
            .Text = ""  '**清空内容
            .Top = Target.Top  '**顶端定位到当前活动单元格顶端
            .Left = Target.Left  '**左边定位到当前活动单元格左边
            .Width = Target.Width '**设定文本框宽为当前活动单元格的宽
            .Height = Target.Height '**设定文本框高为当前活动单元格的高
            .Visible = True  '**显示文本框
        End With
        '**设置listbox
        With Me.infoList
            .Top = Target.Top  '**顶端定位到当前活动单元格顶端
            .Left = Target.Width '**左边定位到当前活动单元格右边
            .Width = Target.Width '**设定列表宽为当前活动单元格的宽
            .Height = 100 '**设定列表框高位100
            .Clear '**清空列表内容
            
            '**将sheet2表A列所有值填入列表框中
            With Sheet2
                For i% = 1 To .Range("a65536").End(xlUp).Row
                    Me.infoList.AddItem .Cells(i, 1)
                Next
            End With
            .Visible = True '**显示列表
        End With
    Else
        Me.txtContext.Visible = False
        Me.infoList.Visible = False
    End If
End Sub